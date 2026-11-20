# Простая запись .xlsx без Excel: листы = @{ Name; Rows (массив массивов); Widths; HeaderRows (индексы строк заголовков, 1-based); Filter (строка заголовка таблицы или 0); Freeze (строк) ; Formats (колонка -> 'int'|'pct'|'money2') }
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
function Esc([string]$s) { [Security.SecurityElement]::Escape(($s -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')) }
function ColName([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][Math]::Floor(($n - 1) / 26) }; $s }
function Write-Xlsx([string]$Path, $Sheets) {
    $inv = [Globalization.CultureInfo]::InvariantCulture
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $zip = [IO.Compression.ZipFile]::Open($Path, 'Create')
    $enc = New-Object Text.UTF8Encoding($false)
    function AddEntry($name, $text) { $e = $zip.CreateEntry($name); $w = New-Object IO.StreamWriter($e.Open(), $enc); $w.Write($text); $w.Close() }
    $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    for ($i = 1; $i -le $Sheets.Count; $i++) { $ct += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>" }
    AddEntry '[Content_Types].xml' ($ct + '</Types>')
    AddEntry '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $wbx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    $defs = ''
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        $wbx += "<sheet name=`"$(Esc $Sheets[$i-1].Name)`" sheetId=`"$i`" r:id=`"rId$i`"/>"
        $rels += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
        $sh = $Sheets[$i-1]
        if ($sh.Filter) { $nc = ($sh.Rows[$sh.Filter - 1]).Count; $defs += "<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$($i-1)`" hidden=`"1`">'$($sh.Name -replace "'", "''")'!`$A`$$($sh.Filter):`$$(ColName $nc)`$$($sh.Rows.Count)</definedName>" }
    }
    $rels += "<Relationship Id=`"rId$($Sheets.Count + 1)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>"
    $wbx += '</sheets>'; if ($defs) { $wbx += "<definedNames>$defs</definedNames>" }; $wbx += '</workbook>'
    AddEntry 'xl/workbook.xml' $wbx
    AddEntry 'xl/_rels/workbook.xml.rels' $rels
    # стили: 0 обычный, 1 заголовок, 2 целое, 3 процент, 4 два знака, 5 жирный, 6 крупный жирный, 7 жирное целое
    AddEntry 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts><fonts count="3"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="13"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E3EF"/></patternFill></fill></fills><borders count="2"><border/><border><bottom style="thin"><color rgb="FF808080"/></bottom></border></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="8"><xf/><xf fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf><xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="9" applyNumberFormat="1"/><xf numFmtId="4" applyNumberFormat="1"/><xf fontId="1" applyFont="1"/><xf fontId="2" applyFont="1"/><xf numFmtId="164" fontId="1" applyNumberFormat="1" applyFont="1"/></cellXfs></styleSheet>'
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        $sh = $Sheets[$i-1]
        $e = $zip.CreateEntry("xl/worksheets/sheet$i.xml"); $w = New-Object IO.StreamWriter($e.Open(), $enc)
        $w.Write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        if ($sh.Freeze) { $w.Write("<sheetViews><sheetView workbookViewId=`"0`"><pane ySplit=`"$($sh.Freeze)`" topLeftCell=`"A$($sh.Freeze + 1)`" activePane=`"bottomLeft`" state=`"frozen`"/></sheetView></sheetViews>") }
        if ($sh.Widths) { $w.Write('<cols>'); for ($j = 0; $j -lt $sh.Widths.Count; $j++) { $w.Write("<col min=`"$($j+1)`" max=`"$($j+1)`" width=`"$($sh.Widths[$j])`" customWidth=`"1`"/>") }; $w.Write('</cols>') }
        $w.Write('<sheetData>')
        $hdr = @{}; foreach ($h in $sh.HeaderRows) { $hdr[[int]$h] = 1 }
        $bold = @{}; foreach ($h in $sh.BoldRows) { $bold[[int]$h] = 1 }
        $big = @{}; foreach ($h in $sh.BigRows) { $big[[int]$h] = 1 }
        $cols = @(); for ($j = 0; $j -lt 60; $j++) { $cols += (ColName ($j + 1)) }
        for ($r = 0; $r -lt $sh.Rows.Count; $r++) {
            $rn = $r + 1; $row = $sh.Rows[$r]
            $ht = if ($hdr[$rn]) { ' ht="45" customHeight="1"' } else { '' }
            $w.Write("<row r=`"$rn`"$ht>")
            for ($j = 0; $j -lt $row.Count; $j++) {
                $v = $row[$j]; if ($null -eq $v -or ($v -is [string] -and $v -eq '')) { continue }
                $ref = $cols[$j] + $rn
                if ($hdr[$rn]) { $st = 1 } elseif ($big[$rn]) { $st = 6 } else { $st = 0 }
                if ($v -is [ValueType]) {
                    if (-not $hdr[$rn]) { $f = $sh.Formats[$j]; $st = if ($f -eq 'pct') { 3 } elseif ($f -eq 'money2') { 4 } elseif ($f -eq 'int') { if ($bold[$rn]) { 7 } else { 2 } } elseif ($bold[$rn]) { 5 } else { 0 } }
                    $w.Write("<c r=`"$ref`" s=`"$st`"><v>$(([double]$v).ToString('R', $inv))</v></c>")
                } else {
                    if ($bold[$rn] -and -not $hdr[$rn]) { $st = 5 }
                    $w.Write("<c r=`"$ref`" s=`"$st`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(Esc ([string]$v))</t></is></c>")
                }
            }
            $w.Write('</row>')
        }
        $w.Write('</sheetData>')
        if ($sh.Filter) { $nc = ($sh.Rows[$sh.Filter - 1]).Count; $w.Write("<autoFilter ref=`"A$($sh.Filter):$(ColName $nc)$($sh.Rows.Count)`"/>") }
        $w.Write('</worksheet>'); $w.Close()
    }
    $zip.Dispose()
}
