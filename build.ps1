Write-Host "Building QuickLoopUp.exe for x64 Windows..."

$webview2Include = ""
$nugetPath = "$env:TEMP\webview2_nuget\extracted\build\native\include"
if (Test-Path "$nugetPath\WebView2.h") {
    $webview2Include = "-I`"$nugetPath`""
    Write-Host "  Found WebView2.h in NuGet cache" -ForegroundColor Cyan
}

# Compile resource (embeds WebView2Loader.dll)
Write-Host "  Compiling resources..." -ForegroundColor Cyan
windres resources.rc -o resources.o
if ($LASTEXITCODE -ne 0) { Write-Host "Resource compilation failed!" -ForegroundColor Red; exit 1 }

$cmd = "g++ main.cpp resources.o -o QuickLoopUp.exe -O3 -mwindows -static -D_WIN32_WINNT=0x0A00 -DNTDDI_VERSION=0x0A000000 $webview2Include -luiautomationcore -lshlwapi -luser32 -lgdi32 -lole32 -loleaut32 -luuid -lwinhttp -lshcore -ldwmapi"
Write-Host "  Compiling..." -ForegroundColor Cyan
Invoke-Expression $cmd

if ($LASTEXITCODE -eq 0) {
    $size = [math]::Round((Get-Item QuickLoopUp.exe).Length / 1KB)
    Write-Host "Build Succeeded: QuickLoopUp.exe ($($size) KB) - single file, no DLL needed!" -ForegroundColor Green
} else {
    Write-Host "Build Failed!" -ForegroundColor Red
}
