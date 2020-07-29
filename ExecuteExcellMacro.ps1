# Excel�̃}�N�������s����

# $ExcelFileName = $Args[0]
$ExcelFileName = "Sample.xlsx"
$SheetName = "Sheet1"
$MacroName = "CreateParam"

$excel = $null
$WorkbookName = $null
$err = 0

try
{
	# �ǂݎ���p��������
	Set-ItemProperty $$ExcelFileNameArgs[0] -Name IsReadOnly -Value $false

	$excel = New-Object -ComObject Excel.Application
	$excel.Visible = $false

	# �x���𖳎�
	$excel.DisplayAlerts = $false

	# Excel�t�@�C�����J��
	$WorkbookName = $excel.Workbooks.Open($ExcelFileName)

	# ActiveSheet��ύX
	#$WorkbookName.Worksheets.item($SheetName).Activate()

	# �}�N�����s
	$excel.Run($MacroName)

	$WorkbookName.Save()
}
catch
{
	# �G���[���b�Z�[�W��\������
	Write-Error("Error"+$_.Exception)
	$err = 1
}
finally
{
	# Excel�����
	$workbook = $null
	$excel.Quit()
	$excel = $null
	[System.GC]::Collect()
}
