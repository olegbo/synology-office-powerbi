# Writes a minimal xlsx: sheet "Продажи" (visible, header + 2 rows), "Скрытый" (hidden).
param([string]$Out)
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
if (Test-Path $Out) { Remove-Item $Out }
$parts = [ordered]@{
'[Content_Types].xml' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>
'@
'_rels/.rels' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
'@
'xl/workbook.xml' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Продажи" sheetId="1" r:id="rId1"/><sheet name="Скрытый" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>
'@
'xl/_rels/workbook.xml.rels' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>
'@
'xl/worksheets/sheet1.xml' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Регион</t></is></c><c r="B1" t="inlineStr"><is><t>Сумма</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>Москва</t></is></c><c r="B2"><v>100</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>Казань</t></is></c><c r="B3"><v>250.5</v></c></row>
</sheetData></worksheet>
'@
'xl/worksheets/sheet2.xml' = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>
'@
}
$zip = [IO.Compression.ZipFile]::Open($Out, 'Create')
foreach ($k in $parts.Keys) {
    $w = New-Object IO.StreamWriter ($zip.CreateEntry($k).Open()), (New-Object Text.UTF8Encoding $false)
    $w.Write($parts[$k]); $w.Dispose()
}
$zip.Dispose()
