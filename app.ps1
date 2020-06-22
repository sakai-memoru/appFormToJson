# param (
#     [Parameter(Mandatory=$true)][string]$excelFile = "appFormToJson.xlsm",
#     [Parameter(Mandatory=$true)][string]$macro = "FieldDefMain.Batch",
#     [Parameter(Mandatory=$true)][string]$formName = "RequestSheet",
#     [Parameter(Mandatory=$true)][boolean]$dumpOn = $false,
#     [Parameter(Mandatory=$true)][boolean]$moveOn = $false
# )

$excelFile = "appFormToJson.xlsm"
$macro = "FormToJsonMain.Batch"
$formName = "RequestSheet"
$dumpOn = $false
$moveOn = $false
##
$curFolder = pwd 
$fullpath = Join-Path $curFolder.Path $excelFile
$excel = new-object -comobject excel.application
$excel.Visible = $false
$workbook = $excel.workbooks.open($fullpath)
$null = $excel.Run($macro, $formName, $dumpOn, $moveOn)
$workbook.close()
$excel.Quit()
echo 'finish ..... !'
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
$null = Remove-Variable excel
