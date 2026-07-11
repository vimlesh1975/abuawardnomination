$ErrorActionPreference = "Stop"

$baseDir = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination"
$pptxPath = Join-Path -Path $baseDir -ChildPath "Vimlesh_GitHub_Projects_Summary.pptx"
$pdfPath = Join-Path -Path $baseDir -ChildPath "Vimlesh_GitHub_Projects_Summary.pdf"

Write-Host "Starting PowerPoint COM Application..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1 

# Open the existing presentation (ReadOnly = $true)
$pres = $ppt.Presentations.Open($pptxPath, -1)

Write-Host "Saving presentation as PDF..."
# 2 = ppFixedFormatTypePDF
$pres.ExportAsFixedFormat($pdfPath, 2)

$pres.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "Created PDF successfully."
