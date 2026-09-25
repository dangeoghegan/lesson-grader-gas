const {test}=require('node:test');
const assert=require('node:assert/strict');
const {context}=require('./helpers.cjs');

test('web access fails closed; Gemini key stored only in Script Properties, never returned',()=>{
  const values={};let email='teacher@school.edu.au';
  const props={getProperty:key=>values[key]||null,setProperty:(key,v)=>{values[key]=v;}};
  const g=context({
    Session:{getActiveUser:()=>({getEmail:()=>email})},
    PropertiesService:{getScriptProperties:()=>props},
    GeminiService:{testGeminiConnection:()=>({success:true,message:'Connected to Gemini successfully.'})}
  }).load('WebAccess.gs');
  assert.equal(g.apiWebAuthStatus().data.authorised,false);
  assert.throws(()=>g.WebAccess.requireTeacher(),/Access denied/);
  values.LESSON_GRADER_ALLOWED_EMAILS='teacher@school.edu.au';
  assert.equal(g.apiWebAuthStatus().data.authorised,true);
  const secret='THIS_IS_A_TEST_KEY_NOT_REAL_123456';
  const result=g.apiWebSetGeminiKey(secret);
  assert.equal(result.success,true);
  assert.equal(values.GEMINI_API_KEY,secret);
  assert.ok(!JSON.stringify(result).includes(secret));
  assert.ok(!JSON.stringify(g.apiWebAuthStatus()).includes(secret));
  assert.equal(g.apiWebTestGeminiKey().success,true);
  email='outsider@school.edu.au';
  assert.throws(()=>g.apiWebSetGeminiKey(secret),/Access denied/);
  email='';assert.equal(g.apiWebAuthStatus().data.authorised,false);
});

test('Gemini transport uses x-goog-api-key header and no key in URL or error response',()=>{
  const key='THIS_IS_A_TEST_KEY_NOT_REAL_123456';
  const calls=[];
  const g=context({
    PropertiesService:{getScriptProperties:()=>({getProperty:()=>key})},
    UrlFetchApp:{fetch:(url,options)=>{
      calls.push({url,options});
      return {getResponseCode:()=>200,getContentText:()=>JSON.stringify({candidates:[{content:{parts:[{text:'{}'}]}}]})};
    }}
  }).load('Code.gs');
  const res=g.GeminiService.callGeminiWithFallback([{role:'user',parts:[{text:'Ping'}]}],false);
  assert.equal(res.success,true);
  assert.equal(calls.length,1);
  assert.equal(calls[0].options.headers['x-goog-api-key'],key);
  assert.ok(!calls[0].url.includes(key));
  assert.ok(!calls[0].url.includes('?key='));
  g.UrlFetchApp.fetch=()=>({getResponseCode:()=>403,getContentText:()=>`Forbidden: ${key}`});
  const failed=g.GeminiService.callGeminiWithFallback([{role:'user',parts:[{text:'Ping'}]}],false);
  assert.equal(failed.success,false);
  assert.ok(!failed.message.includes(key));
});

test('legacy and new public RPCs reject unauthorised callers before Google services run',()=>{
  let usedClassroom=false;
  const mocks={
    WebAccess:{requireTeacher:()=>{throw new Error('Access denied.');}},
    ClassroomService:{listTeacherCourses:()=>{usedClassroom=true;return {success:true};}},
    LessonGraderWeb:{list:()=>{usedClassroom=true;return {success:true};}}
  };
  const g=context(mocks).load('Code.gs','WebAppIntegration.gs','WebGrading.gs');
  assert.throws(()=>g.apiListTeacherCourses(),/Access denied/);
  assert.throws(()=>g.apiWebListCandidateWorkbooks(),/Access denied/);
  assert.throws(()=>g.apiWebEnableGrading('WORKBOOK'),/Access denied/);
  assert.equal(usedClassroom,false);
});
