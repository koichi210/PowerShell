# Excelのマクロを実行する

# $ExcelFileName = $Args[0]
$ExcelFileName = "Sample.xlsx"
$SheetName = "Sheet1"
$MacroName = "CreateParam"

$excel = $null
$WorkbookName = $null
$err = 0

try
{
	# 読み取り専用属性解除
	Set-ItemProperty $$ExcelFileNameArgs[0] -Name IsReadOnly -Value $false

	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false

	# 警告を無視
	$excel.DisplayAlerts = $false

	# Excelファイルを開く
	$WorkbookName = $excel.Workbooks.Open($ExcelFileName)

	# ActiveSheetを変更
	#$WorkbookName.Worksheets.item($SheetName).Activate()

	# マクロ実行
	$excel.Run($MacroName)

	$WorkbookName.Save()
}
catch
{
	# エラーメッセージを表示する
	Write-Error("Error"+$_.Exception)
	$err = 1
}
finally
{
	# Excelを閉じる
	$workbook = $null
	$excel.Quit()
	$excel = $null
	[System.GC]::Collect()
}
