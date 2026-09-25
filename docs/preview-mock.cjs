/* Browser-only design preview. No Google data, no persistence, no real API calls.
   This file is served ONLY by docs/preview.cjs, never deployed to Apps Script. */
(function () {
  'use strict';
  window.__LESSON_GRADER_PREVIEW__ = true;
  var books = [
    { id: 'PREVIEW_WORKBOOK_RUBE_GOLDBERG_2026', name: '7TECHI [2026] - Rube Goldberg Task - Grading' },
    { id: 'PREVIEW_WORKBOOK_WATER_FILTER_2026', name: '8TECHI [2026] - Water Filtration - Grading' }
  ];
  var tasks = { 'PREVIEW_WORKBOOK_RUBE_GOLDBERG_2026':['Rube Goldberg Task'], 'PREVIEW_WORKBOOK_WATER_FILTER_2026':['Water Filtration'] };
  var rubrics = [
    { id:'PREVIEW_RUBRIC_RUBE_GOLDBERG_2026', name:'Rube Goldberg — Design Process Rubric' },
    { id:'PREVIEW_RUBRIC_WATER_FILTER_2026', name:'Water Filtration — Systems Rubric' },
    { id:'PREVIEW_RUBRIC_JEWELLERY_2026', name:'Jewellery Design — Design Folio Rubric' }
  ];
  var names = ['Mia Taylor','Rory Smith','Ava Chen','Noah Wilson','Luca Brown','Harper James'];
  var statuses = ['InReview','New','Approved','New','Locked','InReview'];
  var submissions = [];
  for (var i=0;i<names.length;i++) submissions.push({submissionRecordId:'PREVIEW_SUBMISSION_'+i,studentName:names[i],studentEmail:names[i].toLowerCase().replace(' ','.')+'@school.edu.au',status:statuses[i],currentOfficial:true,version:1});
  var profile={
    schemaVersion:'1.0', profileId:'PREVIEW_PROFILE_RUBE_1',taskName:'Rube Goldberg Task', name:'Rube Goldberg — Design Process Rubric', source:'sharedRubric',
    gradeScaleBands:[{letter:'A',min:85},{letter:'B',min:75},{letter:'C',min:65},{letter:'D',min:50},{letter:'E',min:0}], totalMaxMarks:40,
    criteria:[
      {criterionId:'C01',title:'Investigation & ideas',part:'Design folio',maxMarks:10,outcome:'DT5-2',mdOverride:{id:'C01-MD',text:'Outstanding independent exploration beyond the listed evidence.'},bandOrder:['A','B','C'],bandWeights:{A:1,B:.875,C:.7},bands:{A:[{id:'C01-A-1',text:'Investigates multiple mechanisms with purposeful annotations.'},{id:'C01-A-2',text:'Connects research to informed design decisions.'}],B:[{id:'C01-B-1',text:'Explores several relevant existing mechanisms.'}],C:[{id:'C01-C-1',text:'Shows some relevant research and initial ideas.'}]}},
      {criterionId:'C02',title:'Iteration & planning',part:'Design folio',maxMarks:10,outcome:'DT5-4',mdOverride:{id:'C02-MD',text:'Exceptional refinement through testing and reflection.'},bandOrder:['A','B','C'],bandWeights:{A:1,B:.875,C:.7},bands:{A:[{id:'C02-A-1',text:'Tests, refines and justifies improvements to the design.'},{id:'C02-A-2',text:'Planning identifies safety, resources and realistic milestones.'}],B:[{id:'C02-B-1',text:'Records purposeful design changes.'}],C:[{id:'C02-C-1',text:'Provides a workable initial plan.'}]}},
      {criterionId:'C03',title:'Working prototype',part:'Practical',maxMarks:20,outcome:'DT5-5',mdOverride:{id:'C03-MD',text:'A notably inventive, reliable and carefully finished device.'},bandOrder:['A','B','C'],bandWeights:{A:1,B:.875,C:.7},bands:{A:[{id:'C03-A-1',text:'Chain reaction completes reliably across all stages.'},{id:'C03-A-2',text:'Quality of construction demonstrates precise workmanship.'}],B:[{id:'C03-B-1',text:'Device works with minor guidance or adjustment.'}],C:[{id:'C03-C-1',text:'Most stages function, with some inconsistencies.'}]}}
    ]
  };
  var waterProfile=JSON.parse(JSON.stringify(profile));
  waterProfile.profileId='PREVIEW_PROFILE_WATER_1'; waterProfile.taskName='Water Filtration';
  waterProfile.name='Water Filtration — Systems Rubric';
  waterProfile.criteria[0].title='Research & sustainability';
  waterProfile.criteria[0].bands.A[0].text='Investigates filtration materials and the local water problem.';
  waterProfile.criteria[0].bands.A[1].text='Links material selection to water-quality evidence.';
  waterProfile.criteria[1].title='Testing & iteration';
  waterProfile.criteria[1].bands.A[0].text='Records trials and improves flow and clarity with evidence.';
  waterProfile.criteria[1].bands.A[1].text='Identifies limitations and next steps for safe water treatment.';
  waterProfile.criteria[2].title='Functional filter';
  waterProfile.criteria[2].bands.A[0].text='Prototype demonstrates effective multi-stage filtration.';
  waterProfile.criteria[2].bands.A[1].text='Construction is safe, neat and practical to maintain.';
  waterProfile.criteria[0].mdOverride.text='Exceptional research grounded in environmental evidence.';
  waterProfile.criteria[0].bands.B[0].text='Compares suitable filtration materials.';
  waterProfile.criteria[0].bands.C[0].text='Identifies at least one material and relevant constraint.';
  waterProfile.criteria[1].mdOverride.text='Exceptionally rigorous testing and refinement.';
  waterProfile.criteria[1].bands.B[0].text='Records and acts on test results.';
  waterProfile.criteria[1].bands.C[0].text='Performs a basic filter test.';
  waterProfile.criteria[2].mdOverride.text='Outstanding filtration with thoughtful material use.';
  waterProfile.criteria[2].bands.B[0].text='Filter works with minor issues.';
  waterProfile.criteria[2].bands.C[0].text='Prototype shows partial filtration.';
  function profileForTask(task){return task==='Water Filtration'?waterProfile:profile;}
  var drafts={};
  function copy(o){return JSON.parse(JSON.stringify(o));}
  function ok(data){return {success:true,data:copy(data)};}
  function count(status){return submissions.filter(function(s){return s.status===status;}).length;}
  var methods={
    apiWebAuthStatus:function(){return ok({authorised:true,email:'preview@school.edu.au',keyConfigured:false});},
    apiWebListCandidateWorkbooks:function(){return ok(books);},
    apiWebListSharedRubrics:function(){return ok(rubrics);},
    apiWebListWorkbookTasks:function(bookId){return ok(tasks[bookId]||[]);},
    apiWebGetGradeScale:function(){return ok({weights:{A:1,B:.875,C:.7,D:.575,E:.25},bands:profile.gradeScaleBands});},
    apiWebTaskDashboard:function(bookId,task){return ok({workbook:books.filter(function(b){return b.id===bookId;})[0],taskName:task,prepared:true,profile:{name:profileForTask(task).name,criteriaCount:profileForTask(task).criteria.length,totalMaxMarks:profileForTask(task).totalMaxMarks,observableCount:12},profileVersions:1,stats:{total:submissions.length,approved:count('Approved')+count('Locked'),awaiting:submissions.length-count('Approved')-count('Locked')},submissions:submissions});},
    apiWebGradeDetail:function(_book,_task,id){
      var student=submissions.filter(function(s){return s.submissionRecordId===id;})[0];
      if(!student) return {success:false,message:'Sample student not found.'};
      var d=drafts[id]||{};
      var active=profileForTask(_task);
      var states=d.criteriaMap || (student.status==='Approved' || student.status==='Locked' ? {C01:{checkboxes:{'C01-A-1':true,'C01-A-2':true}},C02:{checkboxes:{'C02-B-1':true}},C03:{checkboxes:{'C03-A-1':true,'C03-A-2':true}}} : {});
      return ok({submission:{submissionRecordId:id,studentName:student.studentName,studentEmail:student.studentEmail,className:_task==='Water Filtration'?'8TECHI':'7TECHI',taskName:active.taskName,status:student.status,version:1,currentOfficial:true,late:false},profile:active,assessmentId:d.assessmentId||(student.status==='Approved'||student.status==='Locked'?'PREVIEW_ASSESSMENT_'+id:''),criteriaMap:states,feedback:d.feedback||{whatWentWell:_task==='Water Filtration'?'The material choices are thoughtful and well justified.':'The mechanism shows careful thinking and a clear sense of purpose.',areasForImprovement:_task==='Water Filtration'?'Run more trials to compare flow and clarity.':'Test the timing of the second stage more consistently.',goalsForNextAssessment:'Record each trial and note what you changed.'},attachments:[{fileName:'design_folio.pdf',mimeType:'application/pdf',link:'https://drive.google.com/',eligible:true},{fileName:'prototype_photos.jpg',mimeType:'image/jpeg',link:'https://drive.google.com/',eligible:true}],history:[{action:'ClassroomImported',when:'2026-09-24',by:'preview@school.edu.au',notes:'Sample data'}],revision:d.revision||'',readOnly:student.status==='Approved'||student.status==='Locked',legacy:false,warning:''});
    },
    apiWebSaveDraft:function(_book,_task,id,criteria,feedback){drafts[id]={criteriaMap:copy(criteria),feedback:copy(feedback),assessmentId:'PREVIEW_ASSESSMENT_'+id,revision:'PREVIEW_REVISION_'+Date.now()};submissions.filter(function(s){return s.submissionRecordId===id;})[0].status='InReview';return {success:true,message:'Sample draft saved in this browser only. No Google data changed.'};},
    apiWebApprove:function(_book,_task,id){submissions.filter(function(s){return s.submissionRecordId===id;})[0].status='Approved';return {success:true,message:'Sample approval shown for this preview. No Google data changed.'};},
    apiWebListTeacherCourses:function(){return {success:true,courses:[{id:'101',name:'7TECHI — Design & Technology'},{id:'202',name:'8TECHI — Systems'}]};},
    apiWebListCourseWork:function(){return {success:true,coursework:[{id:'301',title:'Rube Goldberg Task',dueDate:'2026-10-16'},{id:'302',title:'Research Folio',dueDate:'2026-10-23'}]};},
    apiWebLinkClassroomAssignment:function(){return {success:true,message:'Sample assignment linked in this browser only.'};},
    apiWebPreviewRubric:function(_book,_task,id){var selected=id&&id.indexOf('WATER')!==-1?waterProfile:profile;return ok({rows:selected.criteria.flatMap(function(c){return c.bandOrder.flatMap(function(b){return c.bands[b].map(function(o){return {criterion:c.title,part:c.part,maxMarks:c.maxMarks,outcome:c.outcome,band:b,description:o.text};});});}),preview:selected,validation:{ok:true,errors:[]}});},
    apiWebEnableGrading:function(){return {success:true,message:'This design preview never changes workbooks.'};},
    apiWebImportTask:function(){return {success:true,message:'This design preview never imports real students.'};},
    apiWebAddManualSubmission:function(_book,_task,name,email){var id='PREVIEW_MANUAL_'+submissions.length;submissions.push({submissionRecordId:id,studentName:name,studentEmail:email,status:'New',currentOfficial:true,version:1});return {success:true,submissionRecordId:id,message:'Sample student added in this browser only. No Google data changed.'};},
    apiWebActivateRubric:function(){return {success:false,message:'Design preview: rubric activation is disabled. Use a real school deployment after reviewing the setup guide.'};},
    apiWebCreateWorkbook:function(){return {success:false,message:'Design preview: workbook creation is disabled.'};},
    apiWebRunAi:function(){return {success:false,message:'Design preview: Gemini and student files are not connected.'};},
    apiWebReport:function(){return {success:false,message:'Design preview: PDFs are available from a real deployment.'};},
    apiWebReassess:function(){return {success:false,message:'Design preview: reassessment requires a real deployment.'};}
  };
  function runner(success,failure){
    return new Proxy({}, {get:function(_obj,method){
      if(method==='withSuccessHandler')return function(fn){return runner(fn,failure);};
      if(method==='withFailureHandler')return function(fn){return runner(success,fn);};
      return function(){var args=[].slice.call(arguments);setTimeout(function(){try{if(!methods[method])throw new Error('Unavailable in design preview: '+method);success(methods[method].apply(null,args));}catch(e){if(failure)failure(e);}},210);};
    }});
  }
  window.google={script:{run:runner()}};
}());
