const { test } = require('node:test');
const assert = require('node:assert/strict');
const { context, Workbook, objectRows } = require('./helpers.cjs');
const g = context().load('RubricEngine.gs');
const E=g.RubricEngine;

function makeProfile(task='Rube Goldberg') {
  return E.buildProfileFromRubricRows([
    {criterion:'Design process',part:'Research',maxMarks:8,band:'A',description:'Research informs design choices'},
    {criterion:'Design process',part:'Research',maxMarks:8,band:'A',description:'Ideation is clearly annotated'},
    {criterion:'Design process',part:'Research',maxMarks:8,band:'B',description:'Relevant design research'},
    {criterion:'Practical outcome',part:'Making',maxMarks:12,band:'A',description:'Mechanism works reliably'}
  ],{taskName:task,name:'Rube Goldberg Rubric'});
}

test('imports arbitrary project criteria, parts and mark totals — never jewellery defaults',()=>{
  const p=makeProfile();
  assert.equal(E.validateProfile(p).ok,true);
  assert.equal(p.criteria.length,2);
  assert.equal(p.totalMaxMarks,20);
  assert.deepEqual(Array.from(p.parts),['Research','Making']);
  assert.equal(p.criteria[0].bands.A[0].id,'C01-A-1');
  assert.ok(!E.serializeForAi(p).includes('Jewellery'));
  assert.ok(E.serializeForAi(p).includes('Mechanism works reliably'));
});

test('deterministic scores match legacy weights for a single band; multi-band values cap at full marks',()=>{
  const p=makeProfile();
  const c=p.criteria[0];
  assert.equal(E.scoreCriterion(c,{ 'C01-A-1':true }).points,4);
  const all=E.scoreCriterion(c,{ 'C01-A-1':true,'C01-A-2':true,'C01-B-1':true });
  assert.equal(all.points,8);
  assert.ok(all.rawPoints>all.maxMarks);
  assert.equal(all.grade,'A');
  const md=E.scoreCriterion(c,{'C01-MD':true});
  assert.equal(md.points,8);
  assert.equal(md.score,1);
  const overall=E.scoreProfile(p,{C01:{checkboxes:{'C01-A-1':true}},C02:{checkboxes:{'C02-A-1':true}}});
  assert.equal(overall.totalPoints,16);
  assert.equal(overall.totalMaxMarks,20);
  assert.equal(overall.letter,'B');
});

test('unfamiliar bands require explicit teacher weight; missing marks block activation',()=>{
  const rows=[{criterion:'Water filter',maxMarks:15,band:'LEVEL 4',description:'Clean and clear output'}];
  const unknown=E.buildProfileFromRubricRows(rows,{taskName:'Water filter',name:'Water rubric'});
  assert.equal(E.validateProfile(unknown).ok,false);
  assert.match(E.validateProfile(unknown).errors.join(' '),/explicit 0–1 weight/);
  const reviewed=E.buildProfileFromRubricRows(rows,{taskName:'Water filter',name:'Water rubric',weights:{'LEVEL 4':0.9}});
  assert.equal(E.validateProfile(reviewed).ok,true);
  assert.equal(E.scoreProfile(reviewed,{C01:{checkboxes:{'C01-LEVEL 4-1':true}}}).totalPoints,13.5);
  const missing=E.buildProfileFromRubricRows([{...rows[0],maxMarks:0,band:'A'}],{taskName:'Water filter',name:'Water rubric'});
  assert.equal(E.validateProfile(missing).ok,false);
  assert.throws(()=>E.buildProfileFromRubricRows([{...rows[0],description:''}],{taskName:'Water',name:'Water'}),/needs a criterion, band and observable/);
});

test('migrates a legacy four-column RubricProfiles sheet by header, never overwrites prior JSON',()=>{
  let n=1;
  g.Utilities={getUuid:()=>`PROFILE-${n++}`};
  const ss=new Workbook();
  const original='{ "legacy": true }';
  ss.add('RubricProfiles',[
    ['ProfileID','ProfileName','JSONDefinition','CreatedAt'],
    ['LEGACY','Older draft',original,'2025-01-01']
  ]);
  E.ensureProfilesSheet(ss);
  assert.ok(ss.getSheetByName('RubricProfiles').rows[0].includes('TaskName'));
  const p=makeProfile();
  assert.equal(E.saveProfile(ss,p,'teacher@school.edu.au').success,true);
  assert.equal(E.getActiveProfile(ss,p.taskName).name,p.name);
  const rows=objectRows(ss.getSheetByName('RubricProfiles'));
  assert.equal(rows[0].JSONDefinition,original);
  assert.equal(rows[1].ProfileName,p.name);
  assert.equal(rows[1].TaskName,p.taskName);
  assert.equal(rows[1].Active,true);
  const p2=makeProfile();p2.name='Updated version';
  E.saveProfile(ss,p2,'teacher@school.edu.au');
  const all=objectRows(ss.getSheetByName('RubricProfiles'));
  assert.equal(all[1].Active,false);
  assert.equal(all[2].Active,true);
  assert.equal(E.getProfileById(ss,p.profileId).name,p.name);
  assert.equal(E.getActiveProfile(ss,p.taskName).name,'Updated version');
});

test('report escapes untrusted student text and uses task name from profile',()=>{
  const p=makeProfile();
  const html=E.renderReportHtml(p,{submission:{StudentName:'<script>alert(1)</script>',Class:'7TECHI',SubmissionVersion:1},criteriaMap:{},feedback:{whatWentWell:'<img onerror=alert(1)>'}},null,{taskName:p.taskName});
  assert.ok(html.includes('Rube Goldberg'));
  assert.ok(!html.includes('<script>'));
  assert.ok(!html.includes('<img onerror'));
  assert.ok(html.includes('&lt;script&gt;'));
});

test('same title in two project parts becomes two distinct criteria, not merged',()=>{
  const p=E.buildProfileFromRubricRows([
    {criterion:'Reflection',part:'Design folio',maxMarks:5,band:'A',description:'Explains research decisions'},
    {criterion:'Reflection',part:'Making',maxMarks:10,band:'B',description:'Evaluates prototype testing'}
  ],{taskName:'Independent project',name:'Project rubric'});
  assert.equal(p.criteria.length,2);
  assert.equal(p.criteria[0].part,'Design folio');
  assert.equal(p.criteria[1].part,'Making');
  assert.equal(p.totalMaxMarks,15);
});
