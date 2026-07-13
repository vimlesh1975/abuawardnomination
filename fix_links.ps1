$ErrorActionPreference = "Stop"

$baseDir = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination"
$pptxPath = Join-Path -Path $baseDir -ChildPath "01_OUTSTANDING WORK.pptx"

Write-Host "Starting PowerPoint COM Application..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1 

$pres = $ppt.Presentations.Open($pptxPath, 0)

$changed = $false

foreach ($slide in $pres.Slides) {
    foreach ($shape in $slide.Shapes) {
        if ($shape.HasTextFrame) {
            $textRange = $shape.TextFrame.TextRange
            $text = $textRange.Text
            
            # Find URLs using regex
            $pattern = "https?://github\.com/\S+"
            $matches = [regex]::Matches($text, $pattern)
            
            foreach ($m in $matches) {
                # TextRange.Characters is 1-based index
                $start = $m.Index + 1
                $length = $m.Length
                $url = $m.Value
                
                $linkRange = $textRange.Characters($start, $length)
                
                # Make it a hyperlink
                $link = $linkRange.ActionSettings.Item(1) # ppMouseClick
                $link.Action = 7 # ppActionHyperlink
                $link.Hyperlink.Address = $url
                
                # Make it blue
                $linkRange.Font.Color.RGB = 16711680 # RGB(0,0,255)
                $linkRange.Font.Underline = -1 # msoTrue
                
                Write-Host "Updated link: $url"
                $changed = $true
            }
        }
    }
}

if ($changed) {
    Write-Host "Saving presentation..."
    $pres.Save()
} else {
    Write-Host "No links found."
}

$pres.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "Done."
