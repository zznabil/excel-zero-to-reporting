import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
export const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
export function loadCourse(){
  const context={window:{}};
  for(const file of ['lessons.js','demos.js'])vm.runInNewContext(fs.readFileSync(path.join(root,file),'utf8'),context,{filename:file,timeout:5000});
  return {lessons:context.window.EXCEL_TUTORIAL,demos:context.window.EXCEL_DEMOS};
}
export function plan(){
  const {lessons,demos}=loadCourse();
  return lessons.map(({id,example:e})=>({id,...e,base:demos.calculate(e,e.cells),changed:demos.calculate(e,{...e.cells,[e.change.cell]:e.change.value})}));
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
  fs.mkdirSync(path.join(root,'test-results'),{recursive:true});
  fs.writeFileSync(path.join(root,'test-results','example-plan.json'),JSON.stringify(plan(),null,2));
  console.log('EXAMPLE_PLAN_WRITTEN: '+plan().length+' native comparison cases');
}
