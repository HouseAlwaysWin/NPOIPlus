# NPOIPlus Build Script
# 本地構建和測試腳本

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('Build', 'Test', 'Pack', 'Clean', 'All')]
    [string]$Task = 'All',
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    
    [Parameter(Mandatory=$false)]
    [string]$Version = '1.0.0'
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = $PSScriptRoot
$ProjectFile = Join-Path $ProjectRoot "NPOIPlus\NPOIPlus.csproj"
$TestProject = Join-Path $ProjectRoot "NPOIPlusUnitTest\NPOIPlusUnitTest.csproj"
$OutputDir = Join-Path $ProjectRoot "artifacts"

function Write-TaskHeader {
    param([string]$Message)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

function Clean-Build {
    Write-TaskHeader "清理構建產物 / Cleaning build artifacts"
    
    if (Test-Path $OutputDir) {
        Remove-Item -Path $OutputDir -Recurse -Force
        Write-Host "已刪除 artifacts 目錄" -ForegroundColor Green
    }
    
    dotnet clean $ProjectFile --configuration $Configuration
    dotnet clean $TestProject --configuration $Configuration
    
    Write-Host "清理完成 / Clean completed" -ForegroundColor Green
}

function Build-Project {
    Write-TaskHeader "構建專案 / Building project"
    
    Write-Host "恢復依賴 / Restoring dependencies..." -ForegroundColor Yellow
    dotnet restore $ProjectFile
    
    Write-Host "`n構建 NPOIPlus / Building NPOIPlus..." -ForegroundColor Yellow
    dotnet build $ProjectFile --configuration $Configuration --no-restore
    
    if ($LASTEXITCODE -ne 0) {
        throw "構建失敗 / Build failed"
    }
    
    Write-Host "`n構建完成 / Build completed" -ForegroundColor Green
}

function Test-Project {
    Write-TaskHeader "運行測試 / Running tests"
    
    Write-Host "恢復測試專案依賴 / Restoring test project dependencies..." -ForegroundColor Yellow
    dotnet restore $TestProject
    
    Write-Host "`n構建測試專案 / Building test project..." -ForegroundColor Yellow
    dotnet build $TestProject --configuration $Configuration --no-restore
    
    if ($LASTEXITCODE -ne 0) {
        throw "測試專案構建失敗 / Test project build failed"
    }
    
    Write-Host "`n運行單元測試 / Running unit tests..." -ForegroundColor Yellow
    dotnet test $TestProject `
        --configuration $Configuration `
        --no-build `
        --verbosity normal `
        --logger "console;verbosity=detailed" `
        --collect:"XPlat Code Coverage" `
        --results-directory "./TestResults"
    
    if ($LASTEXITCODE -ne 0) {
        throw "測試失敗 / Tests failed"
    }
    
    Write-Host "`n所有測試通過 / All tests passed" -ForegroundColor Green
}

function Pack-Project {
    Write-TaskHeader "打包 NuGet 套件 / Packing NuGet package"
    
    if (-not (Test-Path $OutputDir)) {
        New-Item -Path $OutputDir -ItemType Directory | Out-Null
    }
    
    Write-Host "版本號 / Version: $Version" -ForegroundColor Yellow
    Write-Host "打包中 / Packing..." -ForegroundColor Yellow
    
    dotnet pack $ProjectFile `
        --configuration $Configuration `
        --no-build `
        --output $OutputDir `
        /p:PackageVersion=$Version `
        /p:Version=$Version
    
    if ($LASTEXITCODE -ne 0) {
        throw "打包失敗 / Pack failed"
    }
    
    Write-Host "`n打包完成 / Pack completed" -ForegroundColor Green
    Write-Host "輸出目錄 / Output directory: $OutputDir" -ForegroundColor Cyan
    Get-ChildItem -Path $OutputDir -Filter *.nupkg | ForEach-Object {
        Write-Host "  - $($_.Name)" -ForegroundColor White
    }
}

function Show-Coverage {
    Write-TaskHeader "生成覆蓋率報告 / Generating coverage report"
    
    $CoverageFiles = Get-ChildItem -Path "./TestResults" -Filter "coverage.cobertura.xml" -Recurse
    
    if ($CoverageFiles.Count -eq 0) {
        Write-Host "未找到覆蓋率文件，請先運行測試 / No coverage files found, please run tests first" -ForegroundColor Yellow
        return
    }
    
    Write-Host "安裝 ReportGenerator / Installing ReportGenerator..." -ForegroundColor Yellow
    dotnet tool install -g dotnet-reportgenerator-globaltool --ignore-failed-sources 2>$null
    
    $ReportDir = Join-Path $ProjectRoot "CoverageReport"
    
    Write-Host "`n生成報告 / Generating report..." -ForegroundColor Yellow
    reportgenerator `
        -reports:"./TestResults/**/coverage.cobertura.xml" `
        -targetdir:$ReportDir `
        -reporttypes:"Html;HtmlSummary"
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "`n覆蓋率報告已生成 / Coverage report generated" -ForegroundColor Green
        Write-Host "報告位置 / Report location: $ReportDir" -ForegroundColor Cyan
        
        $IndexFile = Join-Path $ReportDir "index.html"
        if (Test-Path $IndexFile) {
            Write-Host "`n正在打開報告 / Opening report..." -ForegroundColor Yellow
            Start-Process $IndexFile
        }
    }
}

function Show-Info {
    Write-TaskHeader "環境資訊 / Environment Information"
    
    Write-Host ".NET SDK 版本 / .NET SDK Version:" -ForegroundColor Yellow
    dotnet --version
    
    Write-Host "`n.NET Runtimes:" -ForegroundColor Yellow
    dotnet --list-runtimes
    
    Write-Host "`n專案資訊 / Project Information:" -ForegroundColor Yellow
    Write-Host "  專案文件 / Project file: $ProjectFile" -ForegroundColor White
    Write-Host "  配置 / Configuration: $Configuration" -ForegroundColor White
    Write-Host "  版本 / Version: $Version" -ForegroundColor White
}

# 主程序 / Main
try {
    Show-Info
    
    switch ($Task) {
        'Clean' {
            Clean-Build
        }
        'Build' {
            Build-Project
        }
        'Test' {
            Build-Project
            Test-Project
            Show-Coverage
        }
        'Pack' {
            Build-Project
            Test-Project
            Pack-Project
        }
        'All' {
            Clean-Build
            Build-Project
            Test-Project
            Pack-Project
            Show-Coverage
        }
    }
    
    Write-Host "`n✅ 任務完成 / Task completed successfully" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "`n❌ 錯誤 / Error: $_" -ForegroundColor Red
    exit 1
}

