$projectRoot = (Get-Item $PSScriptRoot).Parent
$src = Get-Item (Join-Path -Path $projectRoot.FullName -ChildPath "src")
$outputFolder = Get-Item (Join-Path -Path $projectRoot.FullName -ChildPath "dev")
$fileName = "BetterArray"
$fileExtension = ".xlsm"
$fullName = $fileName + $fileExtension
$outputPath = Join-Path -Path $outputFolder.FullName -ChildPath $fullName
$backupFolder = Join-Path -Path $outputFolder.FullName -ChildPath "backups"

# Backup old dev file

if (Test-Path -Path $outputPath -PathType Leaf) {
    Write-Host "File [$fullName] exists. Moving to backup"
    try {
        # Create backup folder if not present
        if (-not(Test-Path -Path $backupFolder -PathType Container)) {
            New-Item -ItemType Directory -Path $backupFolder
        }
        $existingFile = Get-Item $outputPath
        $lastModified = $existingFile.LastWriteTime.ToString("yyyy-MM-dd")
        $backUpName = $fileName + "_" + $lastModified + $fileExtension
        $backupPath = Join-Path -Path $backupFolder -ChildPath $backUpName
        $i = 0
        While (Test-Path $backupPath) {
            $i += 1
            $backUpName = "$($fileName)_$($lastModified)($($i))$($fileExtension)"
            $backupPath = Join-Path -Path $backupFolder -ChildPath $backUpName
        }
        $existingFile | Move-Item -Destination $backupPath
        Write-Host "File backed up as '[$backUpName]'."
    }
    catch {
        throw $_.Exception.Message
    }
}

Write-Host "Getting source files"
$macros = Get-ChildItem -Path $src.FullName -Recurse -Include *.bas, *.cls

Write-Host "Creating new excel workbook"
$excel = New-Object -ComObject Excel.Application
$excel.visible = $False
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled
$workbook = $excel.Workbooks.Add()

Write-Host "Importing source code"
try {
    foreach ($macro in $macros) {
        Write-Host "Importing $($macro.FullName)"
        $workbook.VBProject.VBComponents.Import($macro.FullName)
    }
}
catch {
    $msg = "Unable to import source. This error is likely the result of Office Trust Center settings. From within an Office application check Options > Trust Center > Trust Center Settings > Macro Settings > Developer Macro Settings and ensure that 'Trust access to the VBA project object model' is ticked."
    Write-Host $msg -ForegroundColor red
}

# Save and close
$workbook.SaveAs($outputPath, $xlFixedFormat)
$excel.Workbooks.Close()
$excel.Quit()



