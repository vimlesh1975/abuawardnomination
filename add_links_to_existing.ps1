$ErrorActionPreference = "Stop"

$baseDir = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination"
$pptxPath = Join-Path -Path $baseDir -ChildPath "Vimlesh_GitHub_Projects_Summary.pptx"

Write-Host "Starting PowerPoint COM Application..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1 

# Open the existing presentation
$pres = $ppt.Presentations.Open($pptxPath)
$slideWidth = $pres.PageSetup.SlideWidth

# Loop through all slides and add the link to the right of the title
foreach ($slide in $pres.Slides) {
    if ($slide.SlideIndex -ge 2) {
        
        # Cleanup previously added links if they exist, so we don't duplicate
        $shapesToDelete = @()
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame) {
                if ($shape.TextFrame.TextRange.Text -match "example-project-link") {
                    $shapesToDelete += $shape
                }
            }
        }
        foreach ($shape in $shapesToDelete) {
            $shape.Delete()
        }

        # Add Link TextBox on the top-right, aligned to the right
        # Width: 400, Height: 30, Top: 15, Left: $slideWidth - 420
        $linkShape = $slide.Shapes.AddTextbox(1, $slideWidth - 420, 15, 400, 30)
        $linkShape.TextFrame.TextRange.Text = "https://github.com/vimlesh1975/example-project-link"
        $linkShape.TextFrame.TextRange.Font.Size = 14
        $linkShape.TextFrame.TextRange.Font.Color.RGB = 16711680 # Blue color
        $linkShape.TextFrame.TextRange.ParagraphFormat.Alignment = 3 # ppAlignRight
    }
}

Write-Host "Saving presentation..."
$pres.Save()
$pres.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "Updated existing PPTX successfully (Link moved to the right)."
