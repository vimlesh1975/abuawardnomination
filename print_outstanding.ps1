$ErrorActionPreference = "Stop"

$baseDir = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination"
$pptxPath = Join-Path -Path $baseDir -ChildPath "01_OUTSTANDING WORK.pptx"
$pdfPath = Join-Path -Path $baseDir -ChildPath "01_OUTSTANDING WORK.pdf"

if (Test-Path $pdfPath) { Remove-Item $pdfPath -Force }

Write-Host "Starting PowerPoint COM Application..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1 

$pres = $ppt.Presentations.Open($pptxPath, -1)

Write-Host "Printing presentation to PDF..."
try {
    $pres.PrintOptions.ActivePrinter = "Microsoft Print to PDF"
    $pres.PrintOptions.PrintInBackground = 0
    $pres.PrintOptions.FitToPage = -1 # msoTrue
    # PrintOut(From, To, PrintToFile, Copies, Collate)
    $pres.PrintOut(1, $pres.Slides.Count, $pdfPath, 1, 0)
    Write-Host "PrintOut succeeded."
} catch {
    Write-Host "PrintOut failed: $_"
}

$pres.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "Done."
