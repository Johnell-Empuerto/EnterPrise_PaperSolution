$url = "http://localhost:5090/api/form/upload-preview"
$exportUrl = "http://localhost:5090/api/form/save-edited"

function Upload($path) {
    $boundary = [System.Guid]::NewGuid().ToString()
    $fileBytes = [System.IO.File]::ReadAllBytes($path)
    $fileName = Split-Path $path -Leaf
    $header = "--${boundary}`r`nContent-Disposition: form-data; name=`"file`"; filename=`"${fileName}`"`r`nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`r`n`r`n"
    $footer = "`r`n--${boundary}--`r`n"
    $enc = [System.Text.Encoding]::UTF8
    $hBytes = $enc.GetBytes($header)
    $fBytes = $enc.GetBytes($footer)
    $body = New-Object byte[] ($hBytes.Length + $fileBytes.Length + $fBytes.Length)
    [System.Buffer]::BlockCopy($hBytes, 0, $body, 0, $hBytes.Length)
    [System.Buffer]::BlockCopy($fileBytes, 0, $body, $hBytes.Length, $fileBytes.Length)
    [System.Buffer]::BlockCopy($fBytes, 0, $body, $hBytes.Length + $fileBytes.Length, $fBytes.Length)
    return Invoke-RestMethod -Uri $url -Method Post -ContentType "multipart/form-data; boundary=${boundary}" -Body $body
}

$r = Upload "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\FormTest - Copy.xlsx"
$fields = $r.pages[0].fields
$sid = $r.sessionId
$addrs = @('$A$1:$B$2','$C$1:$D$2','$A$3:$D$4','$A$6:$D$7','$A$9:$D$10','$A$12')
$prev = $null
$allOk = $true

for ($cycle = 1; $cycle -le 4; $cycle++) {
    # Build JSON payload
    $json = '{"sessionId":"' + $sid + '","info":{"title":"Test"},"sheets":[{"name":"Sheet1","index":0,"fields":['
    for ($i = 0; $i -lt 6; $i++) {
        if ($i -gt 0) { $json += ',' }
        $json += '{"id":"p1f' + ($i+1) + '","cell":{"address":"' + $addrs[$i] + '","rowIndex":1},"type":0,"value":"",'
        $json += '"placeholder":"F' + ($i+1) + '","defaultValue":"v' + ($i+1) + '","required":true,"maxLength":' + (($i+1)*10) + ','
        $json += '"dataValidation":{"type":"WholeNumber"},"style":{"font":{"name":"Arial","sizePt":' + (11+$i) + ',"colorArgb":"FFFF0000"},"alignment":{"horizontal":"center","vertical":"center"}}}'
    }
    $json += ']}]}'
    
    $resp = Invoke-WebRequest -Uri $exportUrl -Method Post -ContentType "application/json" -Body $json -UseBasicParsing
    $out = "C:\Users\MCF-JOHNELLEEMPUERTO\Documents\test_fc${cycle}.xlsx"
    Set-Content $out $resp.Content -Encoding Byte
    
    $r = Upload $out
    $sid = $r.sessionId
    
    if (-not $r.paperLessConfig) {
        "Cycle ${cycle}: NO PLC!"
        $allOk = $false
        continue
    }
    
    $curr = $r.paperLessConfig.sheets[0].fields | ConvertTo-Json -Depth 5 -Compress
    
    if ($prev) {
        if ($prev -eq $curr) {
            "Cycle ${cycle}: IDENTICAL"
        } else {
            "Cycle ${cycle}: DIFFERENT!"
            $allOk = $false
        }
    } else {
        "Cycle ${cycle}: first"
    }
    
    $prev = $curr
}

if ($allOk) { "PASS: All 4 cycles identical" } else { "FAIL: Some cycles differed" }
