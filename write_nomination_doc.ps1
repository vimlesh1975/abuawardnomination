$ErrorActionPreference = "Stop"
$docxPath = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination\WHY YOU THINK YOUR NOMINEE SHOULD BE SELECTED.docx"

# If the file is 0 bytes, it's not a valid Word file, delete it
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
$selection.Font.Size = 16
$selection.Font.Bold = -1
$selection.TypeText("Justification for the Nomination of Vimlesh Kumar for the ABU Broadcast Engineering Excellence Award")
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.Font.Size = 12
$selection.Font.Bold = 0

$text1 = "Vimlesh Kumar is an exceptionally talented and dedicated broadcast engineer and software developer whose contributions have profoundly impacted the broadcasting industry, both within his organization and across the global broadcast community.`n`n"
$text1 += "His innovative approach to integrating affordable, off-the-shelf hardware (such as Blackmagic DeckLink) with robust open-source software (like CasparCG) has revolutionized cost-effective broadcasting. By developing a comprehensive suite of highly specialized tools, Vimlesh has provided broadcasters with professional-grade capabilities without the exorbitant costs typically associated with proprietary broadcast equipment.`n`n"
$text1 += "Key reasons for his nomination include:`n`n"
$selection.TypeText($text1)

$selection.Font.Bold = -1
$selection.TypeText("1. Development of the ReactCasparClient and WebAnimator: ")
$selection.Font.Bold = 0
$selection.TypeText("He single-handedly developed modern, web-based control clients and animation tools for the widely used CasparCG platform. This enables broadcasters to easily design, schedule, and trigger complex on-air graphics, channel rolls, and lower thirds directly from a web browser.`n`n")

$selection.Font.Bold = -1
$selection.TypeText("2. Hardware Integration Expertise: ")
$selection.Font.Bold = 0
$selection.TypeText("His applications seamlessly integrate with Blackmagic DeckLink hardware to provide advanced capabilities such as multi-channel video and audio recording, slow-motion playback, HTML-to-DeckLink rendering, and off-air logging. His tools are robust enough to serve as primary channel playout video servers in live environments.`n`n")

$selection.Font.Bold = -1
$selection.TypeText("3. Versatile Software Solutions for Live Production: ")
$selection.Font.Bold = 0
$selection.TypeText("Vimlesh has created numerous crucial applications tailored for live broadcast environments, including Web Teleprompters with multi-lingual support, WebSocket inspectors for newsroom systems, Streaming Studio Dashboards, and Election Control Panels capable of handling massive amounts of real-time data.`n`n")

$selection.Font.Bold = -1
$selection.TypeText("4. Global Impact and Open-Source Leadership: ")
$selection.Font.Bold = 0
$selection.TypeText("Instead of keeping his solutions proprietary, Vimlesh has open-sourced an extensive repository of his work (GitHub: vimlesh1975). These applications are actively utilized not only by his parent organization but also by a massive global community of broadcast engineers and CasparCG users. His continuous support, updates, and community engagement demonstrate true leadership.`n`n")

$selection.Font.Bold = -1
$selection.TypeText("5. Technical Mastery and Innovation: ")
$selection.Font.Bold = 0
$selection.TypeText("His repositories showcase a rare and diverse mastery of technologies, spanning HTML, JavaScript, React, Visual Basic .NET, and PowerShell, successfully bridging the gap between web development and low-level broadcast video processing.`n`n")

$text2 = "Vimlesh Kumar embodies the spirit of the ABU Broadcast Engineering Excellence Award. Through his tireless dedication, technical ingenuity, and commitment to the open-source community, he has democratized high-end broadcast capabilities and established new standards for cost-effective media automation."
$selection.TypeText($text2)

# Save
Write-Host "Saving Word Document..."
$doc.SaveAs([ref]$docxPath)
$doc.Close()
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Host "Word Document Created and Saved!"
