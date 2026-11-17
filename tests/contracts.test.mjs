import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {loadCourse,root} from './plan.mjs';
const {lessons,demos}=loadCourse();
const oracle=JSON.parse(fs.readFileSync(path.join(root,'tests/excel-oracle.json'),'utf8'));
const plain=x=>JSON.parse(JSON.stringify(x));
function compare(actual,expected){
  assert.equal(actual.length,expected.length);
  for(let r=0;r<expected.length;r++){assert.equal(actual[r].length,expected[r].length);for(let c=0;c<expected[r].length;c++){
    if(typeof expected[r][c]==='number')assert.ok(typeof actual[r][c]==='number'&&Math.abs(actual[r][c]-expected[r][c])<1e-7,`${actual[r][c]} differs from ${expected[r][c]}`);
    else assert.equal(actual[r][c],expected[r][c]);
  }}
}
const records=file=>{const [header,...rows]=fs.readFileSync(path.join(root,'Data',file),'utf8').trimEnd().split(/\r?\n/);const columns=header.split(',');return rows.map(row=>Object.fromEntries(row.split(',').map((v,i)=>[columns[i],v])));};
const total=(rows,measure)=>rows.reduce((n,r)=>n+measure(r),0);
test('all supported initial and changed demos agree with the observed native Excel oracle',()=>{
  for(const [file,hash]of Object.entries(oracle.sourceHashes))assert.equal(crypto.createHash('sha256').update(fs.readFileSync(path.join(root,file))).digest('hex'),hash,`Native evidence is stale for ${file}; rerun Excel before updating the oracle.`);
  assert.equal(oracle.cases.length,lessons.length);
  for(const l of lessons){const native=oracle.cases.find(c=>c.id===l.id);assert.ok(native?.passed,l.id);compare(demos.calculate(l.example,l.example.cells),native.base);compare(demos.calculate(l.example,{...l.example.cells,[l.example.change.cell]:l.example.change.value}),native.changed);}
});
test('the result verifier rejects a representative plausible but wrong sum',()=>{
  const l=lessons.find(l=>l.id==='s01'),good=plain(demos.calculate(l.example,l.example.cells));compare(good,[[700]]);good[0][0]=0;assert.throws(()=>compare(good,[[700]]));
});
test('criteria limitations stay visible rather than returning misleading zeros',()=>{
  const e=lessons.find(l=>l.id==='s08').example;compare(demos.calculate(e,e.cells),[[1650]]);
  for(const criterion of ['*','I?','~*','>500','=IT'])assert.throws(()=>demos.calculate(e,{...e.cells,D2:criterion}),/literal criteria/);
  compare(demos.calculate(e,{...e.cells,D2:'Absent'}),[[0]]);
});
test('no-match, missing delimiter, non-numeric text and zero denominators remain distinct',()=>{
  const filter=lessons.find(l=>l.id==='s28').example;compare(demos.calculate(filter,{...filter.cells,E2:'Absent'}),[['No matches']]);
  const extraction=lessons.find(l=>l.id==='s33').example;assert.throws(()=>demos.calculate(extraction,{...extraction.cells,A2:'No separator'}),/delimiter/);
  const text=lessons.find(l=>l.id==='s35').example;assert.throws(()=>demos.calculate(text,{...text.cells,A2:'not a number'}),/numeric text/);
  const weighted=lessons.find(l=>l.id==='x06').example;assert.throws(()=>demos.calculate(weighted,{...weighted.cells,B2:0,B3:0}),/positive total/);
  const pmt=lessons.find(l=>l.id==='x08').example;compare(demos.calculate(pmt,{...pmt.cells,B2:0}),[[1000]]);assert.throws(()=>demos.calculate(pmt,{...pmt.cells,C2:0}),/monthly periods/);
});
test('original request and HR project outcomes are independently measured from supplied records',()=>{
  const requests=records('Office-requests.csv');assert.equal(requests.length,8);assert.equal(total(requests,r=>Number(r.Amount)),1170);assert.equal(total(requests.filter(r=>r.Department==='HR'&&r.Status==='Closed'),r=>Number(r.Amount)),320);assert.equal(requests.filter(r=>r.Status==='Closed').length,5);
  const hr=records('Staff-training.csv');assert.equal(hr.length,6);assert.equal(new Set(hr.map(r=>r.EmployeeID)).size,4);assert.equal(total(hr,r=>Number(r.Hours)),26);assert.equal(total(hr.filter(r=>r.Status==='Complete'),r=>Number(r.Hours)),17);assert.equal(hr[0].EmployeeID,'001');
});
test('initial and new monthly files reconcile without duplicate ingestion',()=>{
  const initial=[...records('Monthly-sales/2026-01.csv'),...records('Monthly-sales/2026-02.csv')],updated=[...initial,...records('Add-later/2026-03.csv')];
  assert.equal(initial.length,6);assert.equal(total(initial,r=>Number(r.Units)),11);assert.equal(total(initial,r=>Number(r.Units)*Number(r.UnitPrice)),10200);
  assert.equal(updated.length,8);assert.equal(total(updated,r=>Number(r.Units)),13);assert.equal(total(updated,r=>Number(r.Units)*Number(r.UnitPrice)),12500);
});
test('capstone record grain, quantity, acquisition value and owner exception reconcile',()=>{
  const initial=records('Starter-assets.csv'),updated=[...initial,...records('New-assets.csv')];assert.equal(initial.length,12);assert.equal(total(initial,r=>Number(r.Quantity)),15);assert.equal(total(initial,r=>Number(r.Quantity)*Number(r.UnitCost)),15000);
  assert.equal(updated.length,15);assert.equal(new Set(updated.map(r=>r.AssetID)).size,15);assert.equal(total(updated,r=>Number(r.Quantity)),19);assert.equal(total(updated,r=>Number(r.Quantity)*Number(r.UnitCost)),17700);assert.deepEqual(updated.filter(r=>!r.Owner.trim()).map(r=>r.AssetID),['A015']);
});
