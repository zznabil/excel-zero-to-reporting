import {defineConfig} from '@playwright/test';
const evidence=process.env.COURSE_EVIDENCE||'test-results/browser';
export default defineConfig({
  testDir:'./tests',testMatch:'course.spec.mjs',fullyParallel:false,workers:1,retries:0,
  timeout:45000,expect:{timeout:7000},outputDir:evidence,
  reporter:[['list'],['json',{outputFile:evidence+'-report.json'}]],
  use:{browserName:'chromium',channel:process.env.CI?undefined:'chrome',viewport:{width:1438,height:1000},acceptDownloads:true,screenshot:'only-on-failure',trace:'retain-on-failure'}
});
