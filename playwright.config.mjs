import {defineConfig} from '@playwright/test';
const evidence=process.env.COURSE_EVIDENCE||'test-results/browser';
const courseURL=process.env.COURSE_URL||'http://127.0.0.1:8765/START-HERE.html';
export default defineConfig({
  testDir:'./tests',testMatch:'course.spec.mjs',fullyParallel:false,workers:1,retries:0,
  timeout:45000,expect:{timeout:7000},outputDir:evidence,
  webServer:courseURL.startsWith('http://127.0.0.1:8765/')?{command:'python -m tools.serve --no-browser',cwd:process.env.COURSE_ROOT,url:'http://127.0.0.1:8765/START-HERE.html',reuseExistingServer:false}:undefined,
  reporter:[['list'],['json',{outputFile:evidence+'-report.json'}]],
  use:{browserName:'chromium',channel:process.env.CI?'chromium':'chrome',launchOptions:{ignoreDefaultArgs:['--disable-back-forward-cache']},viewport:{width:1438,height:1000},acceptDownloads:true,screenshot:'only-on-failure',trace:'retain-on-failure'}
});
