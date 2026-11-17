param([switch]$InjectFault)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$run=Join-Path $root ('test-results\projects-'+$(if($InjectFault){'fault-'}else{''})+(Get-Date -Format 'yyyyMMdd-HHmmss'))
[void](New-Item $run -ItemType Directory -Force)
 $app=$null;$book=$null;$recoveryBook=$null;$owned=$false;$checks=[Collections.Generic.List[object]]::new();$failure=$null
function Check($name,$actual,$expected){
 $pass=if($expected -is [string]){[string]$actual -ceq $expected}else{[Math]::Abs([double]$actual-[double]$expected) -lt 0.000001}
 $checks.Add([ordered]@{name=$name;actual=$actual;expected=$expected;passed=$pass})
 if(!$pass){throw "$name differs: $actual != $expected"}
}
function Sheet($name){$s=$book.Worksheets.Add();$s.Name=$name;return $s}
function Formula($s,$address,$formula){$s.Range($address).Formula2=$formula;$s.Calculate();return $s.Range($address).Value2}
function CSVTable($s,$file,$name){
 $records=@(Import-Csv $file);$headers=@($records[0].psobject.Properties.Name)
 for($c=0;$c -lt $headers.Count;$c++){$s.Cells.Item(1,$c+1).Value2=$headers[$c]}
 for($r=0;$r -lt $records.Count;$r++){for($c=0;$c -lt $headers.Count;$c++){$v=$records[$r].($headers[$c]);$cell=$s.Cells.Item($r+2,$c+1);if($headers[$c] -in @('Amount','Hours','Quantity','UnitCost','Units','UnitPrice')){$cell.Value2=[double]$v}else{$cell.NumberFormat='@';$cell.Value2=[string]$v}}}
 $t=$s.ListObjects.Add(1,$s.Range($s.Cells.Item(1,1),$s.Cells.Item($records.Count+1,$headers.Count)),[Type]::Missing,1);$t.Name=$name;return $t
}
function QueryTable($sheet,$name,$m){
 if($m -match '(?i)\b(Web|Sql|OData|AnalysisServices)\.'){throw 'External service query prohibited'}
 [void]$book.Queries.Add($name,$m)
 $connection='OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location='+$name+';Extended Properties=""'
 $t=$sheet.ListObjects.Add(0,$connection,[Type]::Missing,1,$sheet.Range('A1'));$t.Name=$name+'Data'
 $q=$t.QueryTable;$q.CommandType=2;$q.CommandText='SELECT * FROM ['+$name+']';$q.RefreshOnFileOpen=$false;$q.BackgroundQuery=$false
 if(!$q.Refresh($false)){throw 'Query cancelled or failed'}
 return $t
}
function FolderM($folder,$types,$extra){
 $folder=$folder.Replace('"','""')
 return @"
let
 Files=Table.SelectRows(Folder.Files("$folder"),each Text.Lower([Extension])=".csv"),
 Combined=Table.Combine(List.Transform(Files[Content],each Table.PromoteHeaders(Csv.Document(_,[Delimiter=",",Encoding=65001,QuoteStyle=QuoteStyle.Csv]),[PromoteAllScalars=true]))),
 Typed=Table.TransformColumnTypes(Combined,{$types},"en-GB"),
 Result=$extra
in Result
"@
}
function Pivot($source,$name,$measure){
 $s=Sheet ($name+'Report');$cache=$book.PivotCaches().Create(1,$source.Name);$p=$cache.CreatePivotTable($s.Range('A3'),$name+'Pivot');$p.PivotFields('Department').Orientation=1
 [void]$p.AddDataField($p.PivotFields($measure),'Total '+$measure,-4157)
 $p.DataFields().Item(1).NumberFormat='0.00'
 $chart=$s.ChartObjects().Add(240,45,450,300).Chart;$chart.SetSourceData($p.TableRange1);$chart.ChartType=51;$chart.HasTitle=$true;$chart.ChartTitle.Text=if($measure -eq 'Revenue'){'Sales revenue by department (RM)'}else{'Acquisition value by department (RM)'}
 return $p
}
try{
 $app=New-Object -ComObject Excel.Application;if($app.Workbooks.Count -ne 0){throw 'Refusing existing Excel application'};$owned=$true;$app.Visible=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false;$app.AskToUpdateLinks=$false
 $book=$app.Workbooks.Add();$app.Calculation=-4105
 $s=Sheet 'Skill-checks'
 $s.Range('A2').ClearContents();$s.Range('A3').Formula2='=""';$s.Range('A4').Value2=' '
 Check 'truly empty ISBLANK' (Formula $s 'B2' '=ISBLANK(A2)') $true
 Check 'empty-text ISBLANK' (Formula $s 'B3' '=ISBLANK(A3)') $false
 Check 'empty-text equals empty string' (Formula $s 'C3' '=A3=""') $true
 Check 'COUNTBLANK includes empty formula' (Formula $s 'D2' '=COUNTBLANK(A2:A4)') 2
 Check 'COUNTA includes formula and space' (Formula $s 'D3' '=COUNTA(A2:A4)') 2
 $s.Range('F2').Value2=[double]1000;$s.Range('G2').Value2=[double]0.05
 [void]$book.Names.Add('ServiceRate',"='Skill-checks'!`$G`$2")
 Check 'named input' (Formula $s 'H2' '=F2*ServiceRate') 50
 $s.Range('G2').Value2=[double]0.08;$s.Calculate();Check 'changed named input' $s.Range('H2').Value2 80
 $s.Range('F5').Value2='Department';$s.Range('G5').Value2='Amount';$s.Range('F6').Value2='IT';$s.Range('G6').Value2=[double]100;$s.Range('F7').Value2='HR';$s.Range('G7').Value2=[double]50;$s.Range('F8').Value2='IT';$s.Range('G8').Value2=[double]150
 $t=$s.ListObjects.Add(1,$s.Range('F5:G8'),[Type]::Missing,1);$t.Name='VisibleData'
 Check 'SUBTOTAL all visible' (Formula $s 'I6' '=SUBTOTAL(109,G6:G8)') 300
 Check 'SUM all rows' (Formula $s 'I7' '=SUM(G6:G8)') 300
 [void]$t.Range.AutoFilter(1,'IT');$s.Calculate();Check 'SUBTOTAL after filter' $s.Range('I6').Value2 250;Check 'SUM not filtered' $s.Range('I7').Value2 300;$t.AutoFilter.ShowAllData()
 $s.Rows.Item(7).Hidden=$true;$s.Calculate();Check '109 excludes manually hidden' $s.Range('I6').Value2 250
 Check '9 includes manually hidden' (Formula $s 'I8' '=SUBTOTAL(9,G6:G8)') 300;$s.Rows.Item(7).Hidden=$false
 $s.Range('K2').Value2=[double]20;$s.Range('L2').Value2=[double]18;$s.Range('M2').Value2=[double]900
 [void](Formula $s 'N2' '=K2*L2-M2');if(!$s.Range('N2').GoalSeek(0,$s.Range('K2'))){throw 'Goal Seek failed'};Check 'Goal Seek units' $s.Range('K2').Value2 50;Check 'Goal Seek residual' $s.Range('N2').Value2 0
 $s.Range('M2').Value2=[double]990;if(!$s.Range('N2').GoalSeek(0,$s.Range('K2'))){throw 'changed Goal Seek failed'};Check 'changed Goal Seek units' $s.Range('K2').Value2 55
 $s.Activate();$app.ActiveWindow.SplitRow=1;$app.ActiveWindow.SplitColumn=1;$app.ActiveWindow.FreezePanes=$true;Check 'freeze row' $app.ActiveWindow.SplitRow 1;Check 'freeze column' $app.ActiveWindow.SplitColumn 1
 $s.Range('K5:K7').Value2='IT';$validation=$s.Range('L5').Validation;$validation.Add(3,1,1,'IT,HR,FINANCE');$s.Range('L5').Value2='HR';Check 'valid list input' $validation.Value $true;$s.Range('L5').Value2='WRONG';Check 'invalid pasted list input remains detectable' $validation.Value $false
 [void]$s.Range('F6:G8').FormatConditions.Add(2,[Type]::Missing,'=$F6="IT"');Check 'formula conditional-format rule' $s.Range('F6:G8').FormatConditions.Count 1
 $requests=Sheet 'Office-requests';$rt=CSVTable $requests (Join-Path $root 'Data\Office-requests.csv') 'Requests'
 Check 'office rows' $rt.ListRows.Count 8
 Check 'office total' (Formula $requests 'H2' '=SUM(Requests[Amount])') 1170
 Check 'HR closed total' (Formula $requests 'H3' '=SUMIFS(Requests[Amount],Requests[Department],"HR",Requests[Status],"Closed")') 320
 Check 'native wildcard SUMIFS' (Formula $requests 'H4' '=SUMIFS(Requests[Amount],Requests[Department],"*")') 1170
 $hr=Sheet 'Training';$hrTable=CSVTable $hr (Join-Path $root 'Data\Staff-training.csv') 'Bookings'
 Check 'booking rows' $hrTable.ListRows.Count 6
 Check 'booking hours' (Formula $hr 'J2' '=SUM(Bookings[Hours])') 26
 Check 'distinct employees' (Formula $hr 'J3' '=ROWS(UNIQUE(Bookings[EmployeeID]))') 4
 Check 'completed booking hours' (Formula $hr 'J4' '=SUMIFS(Bookings[Hours],Bookings[Status],"Complete")') 17
 Check 'leading-zero employee ID' $hr.Range('B2').Value2 '001'
 $s.Range('Q2').Value2='FIN-2026-014'
 Check 'document prefix' (Formula $s 'R2' '=LEFT(Q2,3)') 'FIN'
 Check 'document year' (Formula $s 'R3' '=MID(Q2,5,4)') '2026'
 Check 'document sequence preserves zero' (Formula $s 'R4' '=RIGHT(Q2,3)') '014'
 Check 'document character count' (Formula $s 'R5' '=LEN(Q2)') 12
 $s.Range('T2').Value2=[double]0;$s.Range('U2').Value2='Review';$s.Range('T3').Value2=[double]0.6;$s.Range('U3').Value2='Pass';$s.Range('T4').Value2=[double]0.8;$s.Range('U4').Value2='Strong'
 foreach($case in @(@('0.59',0.59,'Review'),@('0.6',0.6,'Pass'),@('0.79',0.79,'Pass'),@('0.8',0.8,'Strong'))){
  $s.Range('W2').Value2=[double]$case[1]
  Check ('band '+$case[0]) (Formula $s 'X2' '=XLOOKUP(W2,T2:T4,U2:U4,"No band",-1)') ([string]$case[2])
 }
 $salesFolder=Join-Path $run 'Sales source';[void](New-Item $salesFolder -ItemType Directory);Copy-Item (Join-Path $root 'Data\Monthly-sales\*.csv') $salesFolder
 $salesM=FolderM $salesFolder '{"Date",type date},{"Product",type text},{"Department",type text},{"Units",Int64.Type},{"UnitPrice",type number}' 'Table.AddColumn(Typed,"Revenue",each [Units]*[UnitPrice],type number)'
 $salesSheet=Sheet 'Sales-query';$sales=QueryTable $salesSheet 'Sales' $salesM;$salesPivot=Pivot $sales 'Sales' 'Revenue'
 Check 'initial sales rows' $sales.ListRows.Count 6
 Check 'initial sales units' (Formula $salesSheet 'H2' '=SUM(SalesData[Units])') 11
 Check 'initial sales revenue' (Formula $salesSheet 'H3' '=SUM(SalesData[Revenue])') 10200
 Check 'initial sales Pivot' $salesPivot.GetPivotData('Total Revenue').Value2 10200
 Copy-Item (Join-Path $root 'Data\Add-later\2026-03.csv') $salesFolder
 if(!$sales.QueryTable.Refresh($false)){throw 'Sales refresh failed'};$salesSheet.Calculate()
 Check 'refreshed sales rows' $sales.ListRows.Count 8;Check 'refreshed sales units' $salesSheet.Range('H2').Value2 13;Check 'refreshed sales revenue' $salesSheet.Range('H3').Value2 12500
 $stale=$salesPivot.GetPivotData('Total Revenue').Value2;Check 'stale Pivot demonstrably not new total' ([Math]::Abs($stale-12500) -gt 0.01) $true
 [void]$salesPivot.RefreshTable();Check 'refreshed sales Pivot' $salesPivot.GetPivotData('Total Revenue').Value2 12500
 $book.SaveAs((Join-Path $run 'original-sales-before-recovery.xlsx'),51)
 $copyFolder=Join-Path $run 'Recovery lab Source';[void](New-Item $copyFolder -ItemType Directory);Copy-Item (Join-Path $salesFolder '*.csv') $copyFolder
 $copyFile=Join-Path $run 'My-sales-copy.xlsx';$book.SaveCopyAs($copyFile);$recoveryBook=$app.Workbooks.Open($copyFile,0,$false)
 $copySheet=$recoveryBook.Worksheets.Item('Sales-query');$copyTable=$copySheet.ListObjects.Item('SalesData');$copyPivot=$recoveryBook.Worksheets.Item('SalesReport').PivotTables().Item(1)
 Check 'copied workbook initially still points at original source' ($recoveryBook.Queries.Item('Sales').Formula.Contains($salesFolder)) $true
 $recoveryBook.Queries.Item('Sales').Formula=$salesM.Replace($salesFolder,$copyFolder)
 if(!$copyTable.QueryTable.Refresh($false)){throw 'Disposable source binding failed'};[void]$copyPivot.RefreshTable()
 Check 'rebound copy initial total' $copyPivot.GetPivotData('Total Revenue').Value2 12500
 "Date,Product,Department,Units,UnitPrice`r`n2026-04-01,AP,IT,1,123"|Set-Content (Join-Path $copyFolder 'probe.csv') -Encoding UTF8
 if(!$copyTable.QueryTable.Refresh($false)){throw 'Copy probe refresh failed'};$copySheet.Calculate();[void]$copyPivot.RefreshTable()
 Check 'copy-only probe rows' $copyTable.ListRows.Count 9
 Check 'copy-only probe units' $copySheet.Range('H2').Value2 14
 Check 'copy-only probe total' $copyPivot.GetPivotData('Total Revenue').Value2 12623
 Check 'probe not written into original folder' (Test-Path (Join-Path $salesFolder 'probe.csv')) $false
 if(!$sales.QueryTable.Refresh($false)){throw 'Original source preservation check failed'};$salesSheet.Calculate();[void]$salesPivot.RefreshTable()
 Check 'original stays at eight rows' $sales.ListRows.Count 8
 Check 'original stays at thirteen units' $salesSheet.Range('H2').Value2 13
 Check 'original total unchanged after fresh refresh' $salesPivot.GetPivotData('Total Revenue').Value2 12500
 $moved=Join-Path $run 'Recovery lab Moved-source';Move-Item $copyFolder $moved
 $pathFailed=$false;try{if(!$copyTable.QueryTable.Refresh($false)){$pathFailed=$true}}catch{$pathFailed=$true};Check 'moved disposable source failure detected' $pathFailed $true
 $recoveryBook.Queries.Item('Sales').Formula=$salesM.Replace($salesFolder,$moved)
 if(!$copyTable.QueryTable.Refresh($false)){throw 'Disposable path repair failed'};$copySheet.Calculate();[void]$copyPivot.RefreshTable()
 Check 'repaired copy rows' $copyTable.ListRows.Count 9
 Check 'repaired copy units' $copySheet.Range('H2').Value2 14
 Check 'repaired copy total' $copyPivot.GetPivotData('Total Revenue').Value2 12623
 $recoveryBook.Save();$recoveryBook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($recoveryBook);$recoveryBook=$app.Workbooks.Open($copyFile,0,$true)
 Check 'reopened recovered copy remains independently bound' $recoveryBook.Worksheets.Item('SalesReport').PivotTables().Item(1).GetPivotData('Total Revenue').Value2 12623
 $recoveryBook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($recoveryBook);$recoveryBook=$null
 $assetsFolder=Join-Path $run 'Asset source';[void](New-Item $assetsFolder -ItemType Directory);Copy-Item (Join-Path $root 'Data\Starter-assets.csv') $assetsFolder
 $assetTypes='{"AssetID",type text},{"Department",type text},{"Quantity",Int64.Type},{"UnitCost",type number},{"Owner",type text},{"PurchaseDate",type date},{"RenewalDate",type date}'
 $assetM=FolderM $assetsFolder $assetTypes 'Table.AddColumn(Table.AddColumn(Typed,"TotalValue",each [Quantity]*[UnitCost],type number),"OwnerCheck",each if [Owner]=null or Text.Trim([Owner])="" then "Missing owner" else "OK",type text)'
 $assetSheet=Sheet 'Asset-query';$assets=QueryTable $assetSheet 'Assets' $assetM;$assetPivot=Pivot $assets 'Asset' 'TotalValue'
 Check 'initial asset records' $assets.ListRows.Count 12;Check 'initial devices' (Formula $assetSheet 'M2' '=SUM(AssetsData[Quantity])') 15
 Check 'initial acquisition value' (Formula $assetSheet 'M3' '=SUM(AssetsData[TotalValue])') 15000;Check 'initial asset Pivot' $assetPivot.GetPivotData('Total TotalValue').Value2 15000
 Copy-Item (Join-Path $root 'Data\New-assets.csv') $assetsFolder;if(!$assets.QueryTable.Refresh($false)){throw 'Asset refresh failed'};$assetSheet.Calculate();[void]$assetPivot.RefreshTable()
 Check 'refreshed asset records' $assets.ListRows.Count 15;Check 'refreshed devices' $assetSheet.Range('M2').Value2 19;Check 'refreshed acquisition value' $assetSheet.Range('M3').Value2 17700;Check 'refreshed asset Pivot' $assetPivot.GetPivotData('Total TotalValue').Value2 17700
 Check 'A015 missing owner' (Formula $assetSheet 'M4' '=XLOOKUP("A015",AssetsData[AssetID],AssetsData[OwnerCheck])') 'Missing owner'
 Check 'only one missing owner' (Formula $assetSheet 'M5' '=COUNTIF(AssetsData[OwnerCheck],"Missing owner")') 1
 if($InjectFault){$assetSheet.Range('M3').Formula2='=0';$assetSheet.Calculate();Check 'deliberately broken capstone rejected' $assetSheet.Range('M3').Value2 17700}
 foreach($sheet in $book.Worksheets){[void]$sheet.UsedRange.Columns.AutoFit()}
 $report=$book.Worksheets.Item('AssetReport');$report.Range('A1').Value2='Acquisition value by department (RM)'
 $chartValues=@($report.ChartObjects().Item(1).Chart.SeriesCollection().Item(1).Values)
 Check 'chart has four department values' $chartValues.Count 4
 Check 'chart values reconcile after refresh' (($chartValues|Measure-Object -Sum).Sum) 17700
 $chartObject=$report.ChartObjects().Item(1)
 $report.PageSetup.Orientation=2;$report.PageSetup.Zoom=$false;$report.PageSetup.FitToPagesWide=1;$report.PageSetup.FitToPagesTall=1;$report.PageSetup.PrintArea='$A$1:'+$chartObject.BottomRightCell.Address()
 $printRange=$report.Range($report.PageSetup.PrintArea)
 Check 'print area contains chart right edge' (($printRange.Left+$printRange.Width) -ge ($chartObject.Left+$chartObject.Width)) $true
 Check 'print area contains chart bottom edge' (($printRange.Top+$printRange.Height) -ge ($chartObject.Top+$chartObject.Height)) $true
 $report.ExportAsFixedFormat(0,(Join-Path $run 'asset-report.pdf'))
 $file=Join-Path $run 'original-projects.xlsx';$book.SaveAs($file,51);$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$app.Workbooks.Open($file,0,$true)
 Check 'reopened capstone value' $book.Worksheets.Item('Asset-query').Range('M3').Value2 17700
 Check 'reopened original sales Pivot remains unchanged' $book.Worksheets.Item('SalesReport').PivotTables().Item(1).GetPivotData('Total Revenue').Value2 12500
 Check 'readable report PDF exists' ((Get-Item (Join-Path $run 'asset-report.pdf')).Length -gt 1000) $true
}catch{$failure=$_.Exception.Message}finally{
 $output=[ordered]@{status=$(if($failure){'FAIL'}else{'PASS'});injectedFault=[bool]$InjectFault;checks=$checks;error=$failure;environment='Microsoft 365 desktop Excel on Windows';reopened=(!$failure)}
 $output|ConvertTo-Json -Depth 10|Set-Content (Join-Path $run 'results.json') -Encoding UTF8
 if(!$InjectFault){$output|ConvertTo-Json -Depth 10|Set-Content (Join-Path $root 'test-results\projects-latest.json') -Encoding UTF8}
 if($recoveryBook){$recoveryBook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($recoveryBook)}
 if($book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)};if($owned -and $app){$app.Quit()};if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};[GC]::Collect();[GC]::WaitForPendingFinalizers()
 Write-Output "Evidence: $run"
}
if($failure){throw $failure}
Write-Output 'ORIGINAL_NATIVE_PROJECTS_VERIFIED'
