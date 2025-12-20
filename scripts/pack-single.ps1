# FluentNPOI 單檔合併腳本
# 將所有 FluentNPOI 核心 .cs 檔案合併成單一檔案

param(
    [string]$OutputFile = "FluentNPOI.Single.cs"
)

$ScriptDir = $PSScriptRoot
$ProjectRoot = Split-Path -Parent $ScriptDir
$SourceDir = Join-Path $ProjectRoot "FluentNPOI"
$OutputPath = Join-Path $ProjectRoot $OutputFile

# 定義檔案順序 (依賴順序, ExcelCol 放最後)
$files = @(
    "Models\CellStyleConfig.cs",
    "Streaming\Abstractions\IStreamingRow.cs",
    "Streaming\Abstractions\IRowMapper.cs",
    "Streaming\Mapping\ExcelColumnAttribute.cs",
    "Streaming\Mapping\FluentMapping.cs",
    "Streaming\Mapping\DataTableMapping.cs",
    "Base\FluentCellBase.cs",
    "Base\FluentWorkbookBase.cs",
    "Helpers\FluentMemoryStream.cs",
    "Helpers\FluentNPOIExtensions.cs",
    "Html\HtmlConverter.cs",
    "Stages\FluentCell.cs",
    "Stages\FluentSheet.cs",
    "Stages\FluentTable.cs",
    "Stages\FluentWorkbook.cs",
    "Models\ExcelCol.cs"
)

Write-Host "FluentNPOI Single File Generator" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan

# 收集所有 using 語句
$allUsings = [System.Collections.Generic.HashSet[string]]::new()
$fileContents = @()

foreach ($file in $files) {
    $filePath = Join-Path $SourceDir $file
    if (-not (Test-Path $filePath)) {
        Write-Warning "File not found: $file"
        continue
    }
    
    $lines = Get-Content $filePath -Encoding UTF8
    $contentLines = @()
    $inContent = $false
    
    foreach ($line in $lines) {
        # 只匹配頂層 using 指令 (不匹配 using 語句塊)
        if ($line -match '^\s*using\s+[\w\.]+\s*;\s*$') {
            $using = $line.Trim()
            [void]$allUsings.Add($using)
        } else {
            $contentLines += $line
        }
    }
    
    # 移除開頭的空行
    while ($contentLines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($contentLines[0])) {
        $contentLines = $contentLines[1..($contentLines.Count - 1)]
    }
    
    $fileContents += @{
        Name = $file
        Content = $contentLines -join "`r`n"
    }
    
    Write-Host "  + $file" -ForegroundColor Green
}

# 建立輸出內容
$output = @"
// ============================================================================
// FluentNPOI - Single File Version
// Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
// 
// This file contains all FluentNPOI core functionality in a single file.
// For internal use only.
// 
// GitHub: https://github.com/HouseAlwaysWin/FluentNPOI
// License: MIT
// ============================================================================

#region Using Statements

"@

# 添加排序後的 using 語句
$sortedUsings = $allUsings | Sort-Object
foreach ($using in $sortedUsings) {
    $output += "$using`r`n"
}

$output += @"

#endregion

"@

# 添加每個檔案的內容
foreach ($item in $fileContents) {
    $output += @"

// ============================================================================
// File: $($item.Name)
// ============================================================================

$($item.Content)

"@
}

# 寫入輸出檔案
$output | Out-File -FilePath $OutputPath -Encoding UTF8

$fileInfo = Get-Item $OutputPath
$sizeKB = [math]::Round($fileInfo.Length / 1024, 1)
Write-Host ""
Write-Host "Output: $OutputFile ($sizeKB KB)" -ForegroundColor Cyan
Write-Host "Done!" -ForegroundColor Green
