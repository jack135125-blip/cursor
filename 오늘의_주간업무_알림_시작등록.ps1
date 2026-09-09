# -*- coding: utf-8 -*-
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = Join-Path $root "오늘의_주간업무_알림.py"
if (-not (Test-Path -LiteralPath $py)) {
    throw "프로그램 파일을 찾지 못했습니다: $py"
}

function Resolve-Pythonw {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA "Python\bin\pythonw.exe"),
        (Join-Path $env:LOCALAPPDATA "Python\pythoncore-3.14-64\pythonw.exe")
    )
    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) {
            return $path
        }
    }

    $found = Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA "Python") -Recurse -Filter "pythonw.exe" -ErrorAction SilentlyContinue |
        Where-Object { $_.Length -gt 0 } |
        Select-Object -First 1
    if ($found) {
        return $found.FullName
    }

    throw "pythonw.exe를 찾지 못했습니다. Python이 설치되어 있는지 확인해 주세요."
}

$pythonw = Resolve-Pythonw
$startup = [Environment]::GetFolderPath("Startup")
if (-not $startup) {
    throw "시작 프로그램 폴더를 찾지 못했습니다."
}

Get-ChildItem -LiteralPath $startup -File -ErrorAction SilentlyContinue | Where-Object {
    $_.Extension -in ".bat", ".cmd", ".lnk", ".vbs"
} | ForEach-Object {
    $isOurs = $false
    if ($_.BaseName -match "ChangwonTodaySchedule|주간업무|주간.?업무|알림") {
        $isOurs = $true
    }
    elseif ($_.Extension -in ".bat", ".cmd") {
        $text = Get-Content -LiteralPath $_.FullName -Raw -ErrorAction SilentlyContinue
        if ($text -and ($text -match "오늘의_주간업무_알림|pythonw" -and $text -match "%~dp0")) {
            $isOurs = $true
        }
    }
    if ($isOurs) {
        Remove-Item -LiteralPath $_.FullName -Force
    }
}

$lnkPath = Join-Path $startup "ChangwonTodaySchedule.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($lnkPath)
$shortcut.TargetPath = $pythonw
$shortcut.Arguments = "`"$py`" --startup"
$shortcut.WorkingDirectory = $root
$shortcut.WindowStyle = 7
$shortcut.Description = "오늘의 주간업무 알림"
$shortcut.Save()

Write-Host "시작 프로그램에 등록되었습니다."
Write-Host "바로가기: $lnkPath"
Write-Host "실행 파일: $pythonw"
Write-Host "스크립트: $py"
Write-Host "로그온 후 약 12초 뒤에 알림이 뜹니다."
