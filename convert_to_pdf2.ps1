$ErrorActionPreference = "Stop"

$baseDir = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination"
$pptxPath = Join-Path -Path $baseDir -ChildPath "Vimlesh_GitHub_Projects_Summary.pptx"
$pdfPath = Join-Path -Path $baseDir -ChildPath "Vimlesh_GitHub_Projects_Summary.pdf"

Write-Host "Starting PowerPoint COM Application..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1 

$pres = $ppt.Presentations.Open($pptxPath, -1)

Write-Host "Saving presentation as PDF..."

try {
    # 32 = ppSaveAsPDF
    $pres.SaveAs($pdfPath, 32)
    Write-Host "SaveAs succeeded."
} catch {
    Write-Host "SaveAs failed. Trying ExportAsFixedFormat..."
    try {
        # 2 = ppFixedFormatTypePDF, 1 = ppFixedFormatIntentScreen
        $pres.ExportAsFixedFormat($pdfPath, [System.Object]2, [System.Object]1)
        Write-Host "ExportAsFixedFormat succeeded."
    } catch {
        Write-Host "ExportAsFixedFormat failed: $_"
    }
}

$pres.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
