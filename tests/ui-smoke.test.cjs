const {test}=require('node:test');
const assert=require('node:assert/strict');
const {JSDOM}=require('jsdom');
const fs=require('node:fs');
const path=require('node:path');
const html=fs.readFileSync(path.join(__dirname,'..','WebApp.html'),'utf8');

function preview(designMode=false) {
  const requests=[];
  const bookId='WORKBOOK'+'0'.repeat(30);
  const sourceId='RUBRIC'+'0'.repeat(30);
  const profile={
    profileId:'profile-1',taskName:'Rube Goldberg',name:'Mechanisms & Movement',source:'manual',totalMaxMarks:10,
    gradeScaleBands:[{letter:'A',min:85},{letter:'B',min:75},{letter:'C',min:65},{letter:'D',min:50},{letter:'E',min:0}],
    criteria:[{criterionId:'C01',title:'Functional mechanism',part:'Making',maxMarks:10,outcome:'',mdOverride:{id:'C01-MD',text:'Exceptional performance'},bandOrder:['A'],bandWeights:{A:1},bands:{A:[{id:'C01-A-1',text:'Mechanism works reliably'}]}}]
  };
  let assessmentId='',status='New',check={},revision='';
  const extraStudents=[];
  const responses={
    apiWebAuthStatus:()=>({success:true,data:{authorised:true,email:'dan@school.edu.au',keyConfigured:false}}),
    apiWebListCandidateWorkbooks:()=>({success:true,data:[{id:bookId,name:'7TECHI [2026] - Rube Goldberg - Grading'}]}),
    apiWebListSharedRubrics:()=>({success:true,data:[{id:sourceId,name:'Mechanisms & Movement'}]}),
    apiWebListWorkbookTasks:()=>({success:true,data:['Rube Goldberg']}),
    apiWebGetGradeScale:()=>({success:true,data:{weights:{A:1,B:.875,C:.7,D:.575,E:.25},bands:profile.gradeScaleBands}}),
    apiWebTaskDashboard:()=>({success:true,data:{workbook:{id:bookId,name:'7TECHI [2026] - Rube Goldberg - Grading'},taskName:'Rube Goldberg',prepared:true,profile:{name:'Mechanisms & Movement',criteriaCount:1,totalMaxMarks:10,observableCount:1},profileVersions:1,stats:{total:1,approved:status==='Approved'?1:0,awaiting:status==='Approved'?0:1},submissions:[{submissionRecordId:'SUB-1',studentName:'Mia Taylor',studentEmail:'mia@school.edu.au',status,currentOfficial:true,version:1}].concat(extraStudents)}}),
    apiWebGradeDetail:(_wb,_task,id)=>({success:true,data:{submission:{submissionRecordId:id,studentName:extraStudents.find(s=>s.submissionRecordId===id)?.studentName || 'Mia Taylor',studentEmail:extraStudents.find(s=>s.submissionRecordId===id)?.studentEmail || 'mia@school.edu.au',taskName:'Rube Goldberg',className:'7TECHI',status:id==='SUB-1'?status:'New',version:1,currentOfficial:true},profile,assessmentId:id==='SUB-1'?assessmentId:'',criteriaMap:assessmentId?{C01:{checkboxes:check,teacherWrittenNote:'Well made.'}}:{},feedback:{},attachments:[],history:[],revision,readOnly:status==='Approved',legacy:false,warning:''}}),
    apiWebSaveDraft:(_wb,_task,_sub,criteria)=>{assessmentId='AS-1';check=criteria.C01.checkboxes;status='InReview';revision='R1';return {success:true,message:'Draft saved.',revision};},
    apiWebApprove:()=>{status='Approved';revision='R2';return {success:true,message:'Approved.',revision};},
    apiWebAddManualSubmission:(_wb,_task,name,email)=>{extraStudents.push({submissionRecordId:'SUB-2',studentName:name,studentEmail:email,status:'New',currentOfficial:true,version:1});return {success:true,message:'Added.',submissionRecordId:'SUB-2'};},
    apiWebSetGeminiKey:()=>({success:true,message:'Stored securely.'})
  };
  function runner(success,failure) {
    return new Proxy({}, {get:(_target,prop)=>{
      if (prop==='withSuccessHandler') return handler=>runner(handler,failure);
      if (prop==='withFailureHandler') return handler=>runner(success,handler);
      return (...args)=>queueMicrotask(()=>{
        requests.push([prop,...args]);
        try { if (!responses[prop]) throw new Error(`No mock for ${prop}`); success(responses[prop](...args)); }
        catch(err){ failure(err); }
      });
    }});
  }
  const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://preview.example/',beforeParse(window){
    window.google={script:{run:runner()}};
    window.confirm=()=>true;
    window.scrollTo=()=>{};
    window.HTMLElement.prototype.scrollIntoView=()=>{};
    window.__LESSON_GRADER_PREVIEW__=designMode;
  }});
  return {dom,requests,bookId};
}
async function waitFor(check,what) {
  for(let i=0;i<40;i++){
    if(check())return;
    await new Promise(resolve=>setTimeout(resolve,10));
  }
  throw new Error('Timed out: '+what);
}

test('mobile-ready UI traverses workbook → task → student, saves draft before approval',async()=>{
  const {dom,requests}=preview();
  const doc=dom.window.document;
  await waitFor(()=>doc.querySelector('#overview-list [data-workbook]'), 'workbook list');
  assert.equal(doc.querySelector('#view-overview').hidden,false);
  assert.equal(doc.body.classList.contains('preview-mode'),false);
  assert.ok(html.includes('name="viewport"'));
  assert.ok(html.includes('@media (max-width: 520px)'));
  doc.querySelector('#overview-list [data-workbook]').click();
  await waitFor(()=>doc.querySelector('#task-list [data-task]'),'task list');
  doc.querySelector('#task-list [data-task]').click();
  await waitFor(()=>doc.querySelector('#student-list [data-student]'),'student roster');
  doc.querySelector('#student-list [data-student]').click();
  await waitFor(()=>doc.querySelector('[data-check="C01-A-1"]'),'criterion checkbox');
  assert.equal(doc.querySelector('#student-title').textContent,'Mia Taylor');
  assert.equal(doc.querySelector('#approve-grade').disabled,true);
  const cb=doc.querySelector('[data-check="C01-A-1"]');cb.checked=true;cb.dispatchEvent(new dom.window.Event('change',{bubbles:true}));
  assert.equal(doc.querySelector('#save-state').textContent,'Unsaved changes');
  doc.querySelector('#save-draft').click();
  await waitFor(()=>!doc.querySelector('#approve-grade').disabled,'saved draft');
  assert.equal(doc.querySelector('#live-score').textContent,'10.0');
  doc.querySelector('#approve-grade').click();
  await waitFor(()=>doc.querySelector('#student-status').textContent==='Approved','approved state');
  assert.equal(doc.querySelector('#save-draft').disabled,true);
  assert.equal(doc.querySelector('#lock-grade').hidden,false);
  assert.equal(requests.filter(x=>x[0]==='apiWebSaveDraft').length,1);
  assert.equal(requests.filter(x=>x[0]==='apiWebApprove').length,1);
  dom.window.close();
});

test('Settings saves key without showing it again and manual rubric editor is task-aware',async()=>{
  const {dom,requests}=preview();const doc=dom.window.document;
  await waitFor(()=>doc.querySelector('#overview-list [data-workbook]'),'initial load');
  doc.querySelector('[data-nav="settings"]').click();
  const key='THIS_IS_A_TEST_KEY_NOT_REAL_123456';
  doc.querySelector('#gemini-key').value=key;
  doc.querySelector('#save-key').click();
  await waitFor(()=>doc.querySelector('#key-status').textContent.startsWith('✓ Key configured'),'key status');
  assert.equal(doc.querySelector('#gemini-key').value,'');
  assert.ok(!doc.querySelector('#view-settings').textContent.includes(key));
  assert.equal(requests.filter(x=>x[0]==='apiWebSetGeminiKey').length,1);
  doc.querySelector('#new-manual-rubric').click();
  await waitFor(()=>!doc.querySelector('#view-rubric-editor').hidden,'rubric editor');
  assert.equal(doc.querySelectorAll('.editor-criterion').length,1);
  assert.equal(doc.querySelector('#editor-workbook').options.length,2);
  dom.window.close();
});

test('local design preview labels sample data and never accepts Gemini credentials',async()=>{
  const {dom}=preview(true);const doc=dom.window.document;
  await waitFor(()=>doc.querySelector('#overview-list [data-workbook]'),'initial load');
  assert.equal(doc.body.classList.contains('preview-mode'),true);
  assert.equal(doc.querySelector('#gemini-key').disabled,true);
  assert.equal(doc.querySelector('#save-key').disabled,true);
  assert.equal(doc.querySelector('#upload-rubric').disabled,true);
  dom.window.close();
});

test('manual project submission form adds a student without Classroom setup',async()=>{
  const {dom,requests}=preview();const doc=dom.window.document;
  await waitFor(()=>doc.querySelector('#overview-list [data-workbook]'),'initial load');
  doc.querySelector('#overview-list [data-workbook]').click();
  await waitFor(()=>doc.querySelector('#task-list [data-task]'),'task list');
  doc.querySelector('#task-list [data-task]').click();
  await waitFor(()=>doc.querySelector('#task-add-student')?.disabled===false,'task dashboard');
  doc.querySelector('#task-add-student').click();
  assert.equal(doc.querySelector('#manual-student-form').hidden,false);
  doc.querySelector('#manual-name').value='Casey Lee';
  doc.querySelector('#manual-email').value='casey@school.edu.au';
  doc.querySelector('#manual-student-form').dispatchEvent(new dom.window.Event('submit',{bubbles:true,cancelable:true}));
  await waitFor(()=>doc.querySelector('#student-title').textContent==='Casey Lee','manual student grading view');
  assert.equal(requests.filter(x=>x[0]==='apiWebAddManualSubmission').length,1);
  assert.equal(doc.querySelector('#student-status').textContent,'New');
  dom.window.close();
});
