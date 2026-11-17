param([string]$Only='',[switch]$InjectFault)
$ErrorActionPreference='Stop'
$plan=Get-Content (Join-Path $PSScriptRoot '..\test-results\example-plan.json') -Raw | ConvertFrom-Json
if($Only){$plan=@($plan|Where-Object id -eq $Only);if(!$plan.Count){throw 'Unknown example'}}
$run=Join-Path $PSScriptRoot ('..\test-results\excel-'+$(if($InjectFault){'fault-'}else{''})+(Get-Date -Format 'yyyyMMdd-HHmmss'))
[void](New-Item $run -ItemType Directory)
$app=$null;$book=$null;$owned=$false;$results=[Collections.Generic.List[object]]::new()
function Set-Value($sheet,$address,$value){
 $cell=$sheet.Range($address)
 if($null -eq $value -or ($value -is [string] -and $value -eq '')){[void]$cell.ClearContents()}
 elseif($value -is [string]){$cell.NumberFormat='@';$cell.Value2=$value}
 else{$cell.NumberFormat='General';$cell.Value2=[double]$value}
}
function Matrix($cell){
 if($cell.HasSpill){$range=$cell.SpillingToRange}else{$range=$cell}
 $rows=[Collections.Generic.List[object]]::new()
 for($r=1;$r -le $range.Rows.Count;$r++){
  $row=[Collections.Generic.List[object]]::new()
  for($c=1;$c -le $range.Columns.Count;$c++){ $row.Add($range.Cells.Item($r,$c).Value2) }
  $rows.Add($row.ToArray())
 }
 return ,$rows.ToArray()
}
function Same($actual,$expected){
 if($actual.Count -ne $expected.Count){return $false}
 for($r=0;$r -lt $expected.Count;$r++){
  if($actual[$r].Count -ne $expected[$r].Count){return $false}
  for($c=0;$c -lt $expected[$r].Count;$c++){
   $a=$actual[$r][$c];$e=$expected[$r][$c]
   if($e -is [string]){if($a -isnot [string] -or $a -cne $e){return $false}}
   elseif($a -isnot [double] -and $a -isnot [int]){return $false}
   elseif([Math]::Abs([double]$a-[double]$e) -gt 0.0000001){return $false}
  }
 }
 return $true
}
try{
 $app=New-Object -ComObject Excel.Application
 if($app.Workbooks.Count -ne 0){throw 'Refusing to use an existing application'}
 $owned=$true;$app.Visible=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false;$app.AskToUpdateLinks=$false
 $book=$app.Workbooks.Add();$app.Calculation=-4105
 foreach($item in $plan){
  $sheet=$book.Worksheets.Add();$sheet.Name=$item.id
  foreach($property in $item.cells.psobject.Properties){Set-Value $sheet $property.Name $property.Value}
  if($item.table){$table=$sheet.ListObjects.Add(1,$sheet.Range($item.table),[Type]::Missing,1);$table.Name='Demo_'+$item.id}
  $output=$sheet.Range($item.output);$output.NumberFormat='General'
  $output.Formula2=if($InjectFault){'=999999'}else{$item.formula}
  $sheet.Calculate();$base=Matrix $output
  Set-Value $sheet $item.change.cell $item.change.value
  $sheet.Calculate();$changed=Matrix $output
  $passed=(Same $base $item.base) -and (Same $changed $item.changed)
  $results.Add([ordered]@{id=$item.id;formula=$item.formula;base=$base;changed=$changed;passed=$passed})
  Set-Value $sheet $item.change.cell $item.cells.($item.change.cell)
  [void]$sheet.UsedRange.Columns.AutoFit()
  Write-Output "$($item.id): $passed"
 }
 $book.SaveAs((Join-Path $run 'tutorial-examples.xlsx'),51)
 $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$app.Workbooks.Open((Join-Path $run 'tutorial-examples.xlsx'),0,$true)
 foreach($item in $plan){if(!(Same (Matrix $book.Worksheets.Item($item.id).Range($item.output)) $item.base)){throw "Reopened formula result differs: $($item.id)"}}
 $report=[ordered]@{version=$app.Version;build=$app.Build;injectedFault=[bool]$InjectFault;reopened=$true;results=$results;status=$(if(@($results|Where-Object{!$_.passed}).Count){'FAIL'}else{'PASS'})}
 $report|ConvertTo-Json -Depth 20|Set-Content (Join-Path $run 'results.json') -Encoding UTF8
 $report|ConvertTo-Json -Depth 20|Set-Content (Join-Path $PSScriptRoot '..\test-results\native-latest.json') -Encoding UTF8
 Write-Output "Evidence: $run"
 if($report.status -ne 'PASS'){throw 'TUTORIAL_NATIVE_EXAMPLES_FAILED'}
 Write-Output 'TUTORIAL_NATIVE_EXAMPLES_PASS'
}finally{
 if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($owned -and $app){$app.Quit()}
 if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
