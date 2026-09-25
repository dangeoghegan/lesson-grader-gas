const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context, Workbook, objectRows } = require('./helpers.cjs');

function setup() {
  const ss = new Workbook();
  ss.add('Rube Goldberg',[['Task']]);
  ss.add('Water Filter',[['Task']]);
  // Two web-app-era tables: only minimal headers. prepare() must add missing
  // headers without destroying these rows or changing their existing order.
  ss.add('Submissions',[
    ['SubmissionRecordID','StudentUserID','Status','Task','StudentName','SubmissionVersion','CurrentOfficial'],
    ['SUB-1','STU-1','New','Rube Goldberg','Mia Taylor',1,true]
  ]);
  ss.add('ClassroomConfig',[['ConfigID','CourseID','CourseName','CourseSection','CourseWorkID','AssignmentTitle','MaxPoints','SavedAt','SavedBy','Active','TaskName','DueDate','AutoImported']]);
  ss.add('RubricProfiles',[['ProfileID','ProfileName','JSONDefinition','CreatedAt']]);
  let id=1;
  const mocks={
    Utilities:{getUuid:()=>`id-${id++}`},
    SpreadsheetApp:{flush(){}},
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    WebAccess:{requireTeacher:()=> 'dan@school.edu.au'},
    LessonGraderWeb:{open:(wb)=> {assert.equal(wb,ss.id);return ss;},taskNames:()=>['Rube Goldberg','Water Filter'],isSharedRubric:()=>true},
    DriveService:{isEligibleForAi:(mime)=>mime === 'application/pdf'},
    DriveApp:{getFileById:()=>({makeCopy:()=>({getId:()=> 'backup-1'})})}
  };
  const g=context(mocks).load('RubricEngine.gs','WebGrading.gs');
  return {g,ss};
}
function rowsOne() {return [
  {criterion:'Research',part:'Process',maxMarks:10,band:'A',description:'Evidence of independent research'},
  {criterion:'Research',part:'Process',maxMarks:10,band:'B',description:'Some relevant research'},
  {criterion:'Making',part:'Practical',maxMarks:15,band:'A',description:'Working prototype'}
];}

test('web workflow: prepare, activate own rubric, draft, approve, lock, reassess, pin version',()=>{
  const {g,ss}=setup();const W=g.WebGrading;
  assert.equal(W.listTask(ss.id,'Rube Goldberg').data.prepared,false);
  assert.equal(W.prepare(ss,false).success,true);
  assert.equal(W.listTask(ss.id,'Rube Goldberg').data.prepared,true);
  const saved=W.activateRubric(ss.id,'Rube Goldberg','Rube Goldberg rubric','',rowsOne(),{},true);
  assert.equal(saved.success,true);
  assert.equal(saved.data.totalMaxMarks,25);
  const d=W.detail(ss.id,'Rube Goldberg','SUB-1').data;
  assert.equal(d.profile.criteria.length,2);
  assert.equal(d.assessmentId,'');
  assert.equal(d.revision,'');
  assert.throws(()=>W.approve(ss.id,'Rube Goldberg','SUB-1',''),/Save a complete draft/);
  const draft=W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{
    C01:{checkboxes:{'C01-A-1':true},teacherWrittenNote:'Research observed.'},
    C02:{checkboxes:{'C02-A-1':true},teacherWrittenNote:'Prototype works.'}
  },{whatWentWell:'Good engineering',areasForImprovement:'Measure more',goalsForNextAssessment:'Test early'},'', 'Teacher');
  assert.equal(draft.success,true);
  assert.equal(draft.scores.totalPoints,25);
  assert.equal(objectRows(ss.getSheetByName('CriterionAssessments'))[0].FinalApprovedGrade,'');
  assert.throws(()=>W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{}, {},'outdated','Teacher'),/Reload/);
  const a=W.approve(ss.id,'Rube Goldberg','SUB-1',draft.revision);
  assert.equal(a.success,true);
  assert.equal(a.scores.totalMaxMarks,25);
  assert.equal(objectRows(ss.getSheetByName('ApprovedGrades'))[0].Grade,'A');
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].Status,'Approved');
  assert.throws(()=>W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{}, {},a.revision,'Teacher'),/final/);
  const oldId=d.profile.profileId;
  W.lockAssessment(ss.id,'Rube Goldberg','SUB-1');
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].CurrentOfficial,true);
  W.activateRubric(ss.id,'Rube Goldberg','Newer rubric','',[...rowsOne(),{criterion:'Presentation',maxMarks:5,band:'A',description:'Clear communication'}],{},true);
  assert.equal(W.detail(ss.id,'Rube Goldberg','SUB-1').data.profile.profileId,oldId);
  const r=W.reassess(ss.id,'Rube Goldberg','SUB-1');
  assert.equal(r.success,true);
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].CurrentOfficial,false);
  assert.equal(W.detail(ss.id,'Rube Goldberg',r.submissionRecordId).data.profile.criteria.length,3);
  assert.equal(objectRows(ss.getSheetByName('ApprovedGrades')).length,1);
  assert.throws(()=>W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{}, {},'','Teacher'),/final/);
  assert.equal(W.listTask(ss.id,'Rube Goldberg').data.stats.total,1);
});

test('cannot activate a malformed/unreviewed rubric and cannot grade unknown checkboxes',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  assert.throws(()=>W.activateRubric(ss.id,'Water Filter','Rubric','',rowsOne(),{},false),/Review/);
  assert.throws(()=>W.activateRubric(ss.id,'Water Filter','Rubric','',[{criterion:'Flow',maxMarks:0,band:'A',description:'Works'}],{},true),/mark allocation/);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  assert.throws(()=>W.saveDraft(ss.id,'Rube Goldberg','SUB-1',{C01:{checkboxes:{'C01-FAKE':true}}}, {},'','Teacher'),/Unknown observable/);
  assert.equal(objectRows(ss.getSheetByName('Submissions'))[0].Status,'New');
});

test('task-scoped import keeps two assignments with same student separate and versions changed files',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  const config=ss.getSheetByName('ClassroomConfig');
  config.appendRow(['cfg1','COURSE-1','7TECHI','','WORK-1','Rube Goldberg',100,'','',true,'Rube Goldberg','','']);
  config.appendRow(['cfg2','COURSE-1','7TECHI','','WORK-2','Water Filter',100,'','',true,'Water Filter','','']);
  const payload={
    'WORK-1':[ {id:'CLASS-SUB-1',userId:'STU-2',state:'TURNED_IN',assignmentSubmission:{attachments:[{driveFile:{id:'FILE-A',title:'folio.pdf',mimeType:'application/pdf',alternateLink:'https://drive.google.com/a'}}]}} ],
    'WORK-2':[ {id:'CLASS-SUB-2',userId:'STU-2',state:'TURNED_IN',assignmentSubmission:{attachments:[{driveFile:{id:'FILE-B',title:'prototype.pdf',mimeType:'application/pdf',alternateLink:'https://drive.google.com/b'}}]}} ]
  };
  g.Classroom={Courses:{CourseWork:{StudentSubmissions:{list:(_,work)=>({studentSubmissions:payload[work]})}}},UserProfiles:{get:()=>({name:{fullName:'Rory Smith'},emailAddress:'rory@school.edu.au'})}};
  assert.equal(W.importTask(ss.id,'Rube Goldberg').stats.newCount,1);
  assert.equal(W.importTask(ss.id,'Water Filter').stats.newCount,1);
  let all=objectRows(ss.getSheetByName('Submissions')).filter(s=>s.StudentUserID==='STU-2');
  assert.equal(all.length,2);
  assert.deepEqual(all.map(r=>r.Task).sort(),['Rube Goldberg','Water Filter']);
  assert.equal(W.importTask(ss.id,'Rube Goldberg').stats.unchanged,1);
  payload['WORK-1'][0].assignmentSubmission.attachments[0].driveFile.id='FILE-C';
  assert.equal(W.importTask(ss.id,'Rube Goldberg').stats.newVersions,1);
  all=objectRows(ss.getSheetByName('Submissions')).filter(s=>s.StudentUserID==='STU-2');
  const other=all.find(s=>s.Task==='Water Filter');
  assert.equal(other.AttachmentFileIDsJSON,'["FILE-B"]');
  assert.equal(all.filter(s=>s.Task==='Rube Goldberg' && s.CurrentOfficial===true).length,1);
  assert.equal(all.filter(s=>s.Task==='Rube Goldberg' && s.CurrentOfficial===false).length,1);
});

test('enable-grading backup failure stops all sheet migration',()=>{
  const {g,ss}=setup();const W=g.WebGrading;
  let sawBlank=false;
  g.DriveApp={getFileById:()=>({makeCopy:()=>{
    sawBlank=ss.getSheetByName('ApprovedGrades')===null;
    throw new Error('Drive backup blocked');
  }})};
  assert.throws(()=>W.prepare(ss,true),/Drive backup blocked/);
  assert.equal(sawBlank,true);
  assert.equal(ss.getSheetByName('ApprovedGrades'),null);
  g.DriveApp={getFileById:()=>({makeCopy:()=>({getId:()=> 'BACKUP-ID'})})};
  const enabled=W.prepare(ss,true);
  assert.equal(enabled.backupId,'BACKUP-ID');
  assert.equal(W.listTask(ss.id,'Rube Goldberg').data.prepared,true);
});

test('task-aware AI drafts only valid evidence; logs run without approving or changing Classroom',()=>{
  const {g,ss}=setup();const W=g.WebGrading;
  W.prepare(ss,false);
  W.activateRubric(ss.id,'Rube Goldberg','Rube','',rowsOne(),{},true);
  // Append a file through the prepared sheet's named columns.
  const sh=ss.getSheetByName('SubmissionFiles');
  sh.appendRow(sh.rows[0].map(key=>({SubmissionRecordID:'SUB-1',DriveFileID:'FILE-1',FileName:'folio.pdf',MimeType:'application/pdf'})[key]??''));
  g.DriveService.prepareSingleFileForAi=()=>({status:'reviewed',fileId:'FILE-1',fileName:'folio.pdf',mimeType:'application/pdf',inlineData:{data:'dGVzdA=='}});
  let prompt='';
  g.GeminiService={callGeminiWithFallback:(contents)=>{
    prompt=JSON.stringify(contents);
    return {success:true,modelUsed:'gemini-2.5-flash',text:JSON.stringify({criteria:[
      {criterionId:'C01',tickedCheckboxes:['C01-A-1','C01-MD','C01-FAKE'],evidenceNotes:[{checkboxId:'C01-A-1',fileName:'folio.pdf',note:'Research observed'}]},
      {criterionId:'C02',tickedCheckboxes:['C02-A-1'],evidenceNotes:[]}
    ]})};
  }};
  const result=W.runAi(ss.id,'Rube Goldberg','SUB-1','');
  assert.equal(result.success,true);
  assert.ok(prompt.includes('Research'));
  assert.ok(!prompt.includes('Jewellery Design'));
  const d=W.detail(ss.id,'Rube Goldberg','SUB-1').data;
  assert.equal(d.submission.status,'InReview');
  assert.equal(d.criteriaMap.C01.checkboxes['C01-A-1'],true);
  assert.equal(d.criteriaMap.C01.checkboxes['C01-MD'],undefined);
  assert.equal(d.criteriaMap.C01.checkboxes['C01-FAKE'],undefined);
  assert.equal(d.criteriaMap.C01.aiProposedGrade,'A');
  assert.equal(objectRows(ss.getSheetByName('ApprovedGrades')).length,0);
  assert.equal(objectRows(ss.getSheetByName('AIAssessments')).some(row=>row.RunType==='EvidenceProposal' && row.ModelUsed==='gemini-2.5-flash'),true);
});

test('manual project submission needs no Classroom record and cannot be added twice',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.activateRubric(ss.id,'Water Filter','Water rubric','',rowsOne(),{},true);
  const created=W.addManualSubmission(ss.id,'Water Filter','Casey Lee','casey@school.edu.au','8TECHI');
  assert.equal(created.success,true);
  const d=W.detail(ss.id,'Water Filter',created.submissionRecordId).data;
  assert.equal(d.submission.studentName,'Casey Lee');
  assert.equal(d.submission.className,'8TECHI');
  assert.equal(d.profile.name,'Water rubric');
  const row=objectRows(ss.getSheetByName('Submissions')).find(r=>r.SubmissionRecordID===created.submissionRecordId);
  assert.equal(row.SourceType,'Manual');
  assert.equal(row.ClassroomCourseWorkID,'');
  assert.throws(()=>W.addManualSubmission(ss.id,'Water Filter','Casey Lee','casey@school.edu.au','8TECHI'),/already exists/);
  assert.equal(objectRows(ss.getSheetByName('AssessmentHistory')).find(r=>r.SubmissionRecordID===created.submissionRecordId).Action,'ManualSubmissionCreated');
});

test('Classroom import will not silently duplicate a manually entered student for the same task',()=>{
  const {g,ss}=setup();const W=g.WebGrading;W.prepare(ss,false);
  W.addManualSubmission(ss.id,'Water Filter','Casey Lee','casey@school.edu.au','8TECHI');
  ss.getSheetByName('ClassroomConfig').appendRow(['cfg','COURSE-1','8TECHI','','WORK-2','Water Filter',100,'','',true,'Water Filter','','']);
  g.Classroom={Courses:{CourseWork:{StudentSubmissions:{list:()=>({studentSubmissions:[{
    id:'CLASS-SUB-1',userId:'STU-4',assignmentSubmission:{attachments:[]}
  }]})}}},UserProfiles:{get:()=>({name:{fullName:'Casey Lee'},emailAddress:'casey@school.edu.au'})}};
  assert.throws(()=>W.importTask(ss.id,'Water Filter'),/manual submission already exists/);
  assert.equal(objectRows(ss.getSheetByName('Submissions')).filter(r=>r.Task==='Water Filter').length,1);
});
