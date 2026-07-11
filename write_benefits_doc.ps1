$ErrorActionPreference = "Stop"
$docxPath = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination\BENEFITS ACCRUED TO YOUR ORGANISATION.docx"

# If the file exists and is 0 bytes, delete it
if (Test-Path $docxPath) {
    $file = Get-Item $docxPath
    if ($file.Length -eq 0) {
        Remove-Item $docxPath -Force
    }
}

Write-Host "Starting Word Application..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false

$doc = $word.Documents.Add()
$selection = $word.Selection

$selection.Font.Name = "Calibri"
$selection.Font.Size = 18
$selection.Font.Bold = -1
$selection.TypeText("Benefits Accrued to the Organization")
$selection.TypeParagraph()
$selection.TypeParagraph()

# Function to write a section
function Write-Section {
    param([string]$Title, [string]$Content)
    
    $selection.Font.Size = 14
    $selection.Font.Bold = -1
    $selection.TypeText($Title)
    $selection.TypeParagraph()
    
    $selection.Font.Size = 12
    $selection.Font.Bold = 0
    $selection.TypeText($Content)
    $selection.TypeParagraph()
    $selection.TypeParagraph()
}

$savings = "By developing custom in-house software solutions using affordable commercial-off-the-shelf (COTS) hardware (like Blackmagic DeckLink cards) and open-source software (CasparCG), the organization has saved significantly on expensive proprietary broadcast equipment and recurring licensing fees. Tools such as the multi-channel video and audio recorders effectively replace costly dedicated hardware, allowing the organization to operate robustly on a fraction of a typical broadcast budget."
Write-Section -Title "Operational savings" -Content $savings

$services = "The creation of tools like the Web Teleprompter with multi-lingual support, the CasparCG Election Control Panel, and the Streaming Studio Dashboard has empowered the organization to launch dynamic new live broadcasts. These tools provide advanced, specialized capabilities - such as real-time election data processing and remote web-based graphics scheduling - that were previously unattainable without purchasing entirely new, specialized turnkey systems."
Write-Section -Title "Enabling new services" -Content $services

$resource = "Web-based controllers and schedulers (such as the ReactCasparClient and News Scroll Scheduler) have decoupled operators from physical broadcast hardware. Staff can now manage rundowns, schedule advertisements, and trigger complex graphics over the local network from any standard web browser. This flexibility drastically improves personnel deployment, reduces training times, and allows a single operator to efficiently manage multiple broadcast channels simultaneously."
Write-Section -Title "Improved resource management" -Content $resource

$research = "The continued development and integration of these tools represent ongoing internal research into modernizing broadcast workflows. By bridging modern web technologies (React, WebSockets, HTML5) with low-level SDI broadcast hardware, the organization has created a flexible testbed for future-proofing its broadcast infrastructure and exploring next-generation, cloud-ready media playout solutions."
Write-Section -Title "Research activities" -Content $research

$others = "By open-sourcing many of these tools, the organization has established itself as a valuable contributor to the global broadcast engineering community. The extensive use of these applications worldwide serves as a testament to the high quality and reliability of the internal engineering efforts, ultimately elevating the technical reputation of the organization on an international stage."
Write-Section -Title "Others" -Content $others

# Save
Write-Host "Saving Word Document..."
$doc.SaveAs([ref]$docxPath)
$doc.Close()
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Host "Word Document Created and Saved!"
