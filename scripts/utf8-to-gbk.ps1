# ============================================================
# utf8-to-gbk.ps1 -- Convert UTF-8 files (with or without BOM)
#                     to GBK/ANSI (code page 936) for VB6 sources.
#
# VB6 reads source files as ANSI; UTF-8 (especially with a BOM)
# garbles Chinese comments/strings. Run this after editing
# .bas/.cls/.frm with tools that save UTF-8.
#
# Detection logic:
#   - UTF-8 BOM (EF BB BF)  -> convert (BOM stripped)
#   - valid UTF-8 w/ non-ASCII -> convert
#   - pure ASCII            -> skip (identical in both encodings)
#   - invalid UTF-8         -> skip (assume already GBK/ANSI)
#
# Usage:
#   .\scripts\utf8-to-gbk.ps1                       # project dir, VB6 extensions
#   .\scripts\utf8-to-gbk.ps1 -Path src -Recurse
#   .\scripts\utf8-to-gbk.ps1 -Path file.bas
#   .\scripts\utf8-to-gbk.ps1 -Include *.bas,*.frm
# ============================================================
param(
    [string]$Path = ".",
    [switch]$Recurse,
    [string[]]$Include = @("*.bas", "*.cls", "*.frm", "*.vbp")
)

$ErrorActionPreference = "Stop"
$gbk  = [System.Text.Encoding]::GetEncoding(936)
$utf8 = New-Object System.Text.UTF8Encoding($false)

function Test-Utf8Bytes([byte[]]$bytes) {
    $i = 0
    $n = $bytes.Length
    while ($i -lt $n) {
        $b = $bytes[$i]
        if ($b -le 0x7F) {
            $i += 1; continue
        } elseif ($b -ge 0xC2 -and $b -le 0xDF) {
            if ($i + 1 -ge $n -or $bytes[$i+1] -lt 0x80 -or $bytes[$i+1] -gt 0xBF) { return $false }
            $i += 2
        } elseif ($b -ge 0xE0 -and $b -le 0xEF) {
            if ($i + 2 -ge $n) { return $false }
            $b1 = $bytes[$i+1]; $b2 = $bytes[$i+2]
            if ($b1 -lt 0x80 -or $b1 -gt 0xBF -or $b2 -lt 0x80 -or $b2 -gt 0xBF) { return $false }
            if ($b -eq 0xE0 -and $b1 -lt 0xA0) { return $false }  # overlong
            if ($b -eq 0xED -and $b1 -ge 0xA0) { return $false }  # surrogate
            $i += 3
        } elseif ($b -ge 0xF0 -and $b -le 0xF4) {
            if ($i + 3 -ge $n) { return $false }
            foreach ($k in 1..3) {
                $c = $bytes[$i+$k]
                if ($c -lt 0x80 -or $c -gt 0xBF) { return $false }
            }
            $i += 4
        } else {
            return $false
        }
    }
    return $true
}

# ---- collect target files ----
if (Test-Path $Path -PathType Container) {
    $files = Get-ChildItem $Path -File -Recurse:$Recurse | Where-Object {
        $name = $_.Name
        $Include | Where-Object { $name -like $_ }
    }
} elseif (Test-Path $Path -PathType Leaf) {
    $files = @(Get-Item $Path)
} else {
    Write-Error "Path not found: $Path"
    exit 1
}

if (-not $files) { Write-Host "No matching files."; exit 0 }

# ---- convert ----
$converted = 0; $skipped = 0
foreach ($f in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    if ($hasBom) {
        $text = $utf8.GetString($bytes, 3, $bytes.Length - 3)
        [System.IO.File]::WriteAllBytes($f.FullName, $gbk.GetBytes($text))
        Write-Host "  CONVERTED: $($f.FullName)  (UTF-8 BOM -> GBK)"
        $converted++
    } elseif (Test-Utf8Bytes $bytes) {
        if (($bytes | Where-Object { $_ -gt 0x7F } | Select-Object -First 1)) {
            $text = $utf8.GetString($bytes)
            [System.IO.File]::WriteAllBytes($f.FullName, $gbk.GetBytes($text))
            Write-Host "  CONVERTED: $($f.FullName)  (UTF-8 -> GBK)"
            $converted++
        } else {
            Write-Host "  SKIP: $($f.Name)  (pure ASCII)"
            $skipped++
        }
    } else {
        Write-Host "  SKIP: $($f.Name)  (already GBK/ANSI)"
        $skipped++
    }
}
Write-Host ""
Write-Host "Done: $converted converted, $skipped skipped."
