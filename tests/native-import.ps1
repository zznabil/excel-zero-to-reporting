param([Parameter(Mandatory=$true)][string]$Workbook)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$file=[IO.Path]::GetFullPath($Workbook)
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'test-results'))+'\'
if(!$file.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($file) -notmatch '^ui-training-import-\d+\.xlsx$'){throw 'Only the newly GUI-authored QA import is authorised'}
$app=$null;$book=$null;$owned=$false;$errorText=$null;$checks=[Collections.Generic.List[object]]::new()
function Check($name,$condition){$checks.Add([ordered]@{name=$name;passed=[bool]$condition});if(!$condition){throw $name}}
function Preserved($sheet){$v=$sheet.Range('B2').Value2;return ($v -is [string] -and $v -ceq '001')}
function CheckIDs($sheet,$phase){
 $expected=@('001','002','003','001','004','002')
 for($i=0;$i -lt $expected.Count;$i++){$v=$sheet.Cells.Item($i+2,2).Value2;Check "$phase text ID row $($i+2)" ($v -is [string] -and $v -ceq $expected[$i])}
 Check "$phase first start date is numeric 5 January" ($sheet.Range('G2').Value2 -is [double] -and [Math]::Abs($sheet.Range('G2').Value2-([datetime]'2026-01-05').ToOADate()) -lt 0.000001)
}
try{
 $app=New-Object -ComObject Excel.Application;if($app.Workbooks.Count -ne 0){throw 'Refusing existing application'};$owned=$true;$app.Visible=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false;$app.AskToUpdateLinks=$false
 $book=$app.Workbooks.Open($file,0,$true);$sheet=$book.Worksheets.Item('Staff-training');CheckIDs $sheet 'GUI saved load'
 $table=$sheet.ListObjects.Item(1);Check 'actual loaded Table contains six records' ($table.ListRows.Count -eq 6)
 $table.QueryTable.BackgroundQuery=$false;if(!$table.QueryTable.Refresh($false)){throw 'GUI-authored query refresh failed'};CheckIDs $sheet 'refreshed'
 $control=$book.Worksheets.Add();$control.Name='Wrong-type-control';$control.Range('B2').NumberFormat='000';$control.Range('B2').Value2=[double]1
 Check 'numeric 1 can look like 001' ($control.Range('B2').Text -ceq '001')
 Check 'numeric 1 is rejected despite identical display' (!(Preserved $control))
 $control.Range('B2').NumberFormat='@';$control.Range('B2').Value2='1';Check 'stripped then converted text 1 is rejected' (!(Preserved $control))
 $copy=Join-Path $root ('test-results\ui-import-verified-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.xlsx');$book.SaveCopyAs($copy);$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$app.Workbooks.Open($copy,0,$true);CheckIDs $book.Worksheets.Item('Staff-training') 'reopened'
}catch{$errorText=$_.Exception.Message}finally{
 $report=[ordered]@{status=$(if($errorText){'FAIL'}else{'PASS'});origin='Real desktop Excel GUI: From Text/CSV, Transform Data, delete automatic Changed Type, set Text/numeric/date types, Close and Load to Table, Refresh All, native Save As';checks=$checks;error=$errorText}
 $report|ConvertTo-Json -Depth 8|Set-Content (Join-Path $root 'test-results\ui-import-results.json') -Encoding UTF8
 if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)};if($owned -and $app){$app.Quit()};if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};[GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if($errorText){throw $errorText}
Write-Output 'GUI_IMPORTED_TEXT_IDS_VERIFIED'
