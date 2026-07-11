$ErrorActionPreference = "Stop"

$baseDir = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination"
$pptxPath = Join-Path -Path $baseDir -ChildPath "Vimlesh_GitHub_Projects_Summary.pptx"

Write-Host "Starting PowerPoint COM Application..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1 

# Open the existing presentation
$pres = $ppt.Presentations.Open($pptxPath)

# Loop through all slides and find the link text boxes
foreach ($slide in $pres.Slides) {
    if ($slide.SlideIndex -ge 2) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame) {
                if ($shape.TextFrame.HasText) {
                    $text = $shape.TextFrame.TextRange.Text
                    if ($text -match "^https?://") {
                        # Add a real hyperlink to the text so it's clickable
                        # 1 = ppMouseClick
                        $shape.TextFrame.TextRange.ActionSettings.Item(1).Hyperlink.Address = $text.Trim()
                    }
                }
            }
        }
    }
}

Write-Host "Saving presentation..."
$pres.Save()
$pres.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "Updated existing PPTX successfully (Links made clickable)."
