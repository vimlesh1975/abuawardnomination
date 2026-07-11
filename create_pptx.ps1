$ErrorActionPreference = "Stop"

$data = @(
    @{ Title="DeckLink Scroll Player"; Desc="UI for a vertical and horizontal scroll player with DeckLink device selection, text formatting options, and scrolling text."; Image="decklink_vertical_scroll_player.png" },
    @{ Title="News Scroll Scheduler for R3 Space engine Wtvision"; Desc="UI for scheduling news scrolls with tabs for various templates, displaying multilingual text in a grid."; Image="casparcg_news_scroll.png" },
    @{ Title="DeckLink Audio Recorder"; Desc="UI for audio recording, showing hardware configuration and live audio level meters for 16 channels."; Image="decklink_audio_recorder.png" },
    @{ Title="CasparCG Media Playback"; Desc="CasparCG media playback interface with playlist of videos and server connection status."; Image="caspar_media_playback.png" },
    @{ Title="DecklinkPlayer"; Desc="UI showing DecklinkPlayer playlist of video files, and playback controls.It also includes slow and fast motio playbacks, reverse playback. It can act as channel playout video server."; Image="decklink_player.png" },
    @{ Title="Web Teleprompter Interface"; Desc="Web-based teleprompter interface displaying Hindi text with speed controls and reading options. It can connect to database , take word or notepad. It can work with Shuttle pro v2, and RF remote conrol, Keybord control. It can be operated from mobile also."; Image="web_teleprompter.png" },
    @{ Title="Samvad WebSocket Inspector"; Desc="Web UI for managing stories and run orders for a teleprompter, showing Hindi text content. It is separate interface than Samvad provides. It works with shuttle pro v2."; Image="samvad_websocket_inspector.png" },
    @{ Title="4 Channel DeckLink Recorder"; Desc="UI for a 4-channel video recording application, showing 4 simultaneous video inputs from DeckLink Duo cards. One channel player is also there. Many recording profiles are included."; Image="4_channel_decklink_recorder.png" },
    @{ Title="CefToDecklink HTML Renderer"; Desc="UI for HTML to DeckLink Renderer, showing HTML URL pointing to a URL and a live preview."; Image="cef_to_decklink.png" },
    @{ Title="CasparCG Scroll for advertisement"; Desc="Web UI for scheduling scrolling text on CasparCG, showing fields for host, font family, layer, and speed."; Image="casparcg_scroll_scheduler.png" },
    @{ Title="Streaming Studio Dashboard"; Desc="Web-based studio interface showing live guest video feeds and invite link generator."; Image="webrtc_studio_getstreamio.png" },
    @{ Title="Audio Off-Air Logger"; Desc="UI showing audio input device, waveform, and a list of saved audio clips with download options."; Image="audio_off_air_logger.png" },
    @{ Title="ReactCasparClient"; Desc="A web-based client for CasparCG showing designer tools and a live output preview."; Image="react_caspar_client_designer.png" },
    @{ Title="WebAnimator for CasparCG"; Desc="Web-based animation tool for CasparCG graphics showing a timeline, keyframes, and a preview of a sports graphic."; Image="react_caspar_client_webanimator.png" },
    @{ Title="CasparCG Election Control Panel"; Desc="Web dashboard for controlling election graphics, showing party data and seats controls."; Image="casparcg_election_control_panel.png" },
    @{ Title="Media Explorer and Screen Consumer"; Desc="UI showing a media explorer file tree and a grid interface for playing video blocks."; Image="project_screenshot_16.png" },
    @{ Title="Playout Counter"; Desc="Application showing countdown timers for News, Commercial, and Promo playout segments."; Image="project_screenshot_17.png" },
    @{ Title="DD Chennai Quiz Graphics"; Desc="Web dashboard for triggering quiz graphics showing a gallery of images and video files."; Image="project_screenshot_18.png" },
    @{ Title="Web Audio Recorder"; Desc="Simple web-based audio recorder showing a live waveform and a list of past recordings."; Image="project_screenshot_19.png" },
    @{ Title="Election Data Spreadsheet Dashboard"; Desc="UI showing tables with municipal election data, listing parties and vote counts for various regions."; Image="project_screenshot_20.png" },
    @{ Title="Election Results Graphic Preview"; Desc="A full-screen graphic displaying Marathi text and tables with election results for various regions."; Image="project_screenshot_21.png" }
)

Write-Host "Starting PowerPoint COM Application..."
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1 

$pres = $ppt.Presentations.Add()
$pres.PageSetup.SlideWidth = 1440
$pres.PageSetup.SlideHeight = 810
$slideWidth = $pres.PageSetup.SlideWidth
$slideHeight = $pres.PageSetup.SlideHeight

$baseDir = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination"

    $introSlide = $pres.Slides.Add(1, 12)
    $introTitle = $introSlide.Shapes.AddTextbox(1, 40, 40, $slideWidth - 80, 80)
    $introTitle.TextFrame.TextRange.Text = "ABU Broadcast Engineering Excellence Award`r`nIntroduction"
    $introTitle.TextFrame.TextRange.Font.Size = 32
    $introTitle.TextFrame.TextRange.Font.Bold = -1
    $introTitle.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    
    $introDesc = $introSlide.Shapes.AddTextbox(1, 40, 140, $slideWidth - 80, $slideHeight - 160)
    $introDesc.TextFrame.TextRange.Text = "The ABU Broadcast Engineering Excellence Award is open to staff or former staff of ABU broadcaster members. It recognises contributions of an outstanding nature in broadcast engineering and related disciplines.`r`n`r`nVimlesh Kumar (GitHub: vimlesh1975) is a developer specializing in broadcast graphics, media automation, and hardware integration. His work features robust solutions for live broadcasting environments using CasparCG and Blackmagic DeckLink hardware. His projects are mostly open for public and being used by parent organization as well globally. He has a huge application based on Casparcg and being used by CasparCG community.`r`n`r`nHis repositories highlight expertise across multiple languages including JavaScript, HTML, and Visual Basic .NET. Notable projects include web-based clients for CasparCG (ReactCasparClient), HTML-to-DeckLink rendering tools (Ceftodecklink), multi-channel video and audio recorders, Channel playout - casparmediaplayback, Game Graphics applications, Channel rolls, scrolls, Breaking news, logo, lower third, general graphics, streaming interface, off air logger, trimmer, transcoder, silence detector and auto changeover, slow motion player, Xdcam playou server, vdcp controller, RS422 VTR controller.`r`n`r`nThis document summarizes several of his key projects, featuring user interfaces that streamline broadcast operations and scheduling."
    $introDesc.TextFrame.TextRange.Font.Size = 20
    $introDesc.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    $introDesc.TextFrame.WordWrap = -1

foreach ($item in $data) {
    # 12 = ppLayoutBlank
    $slide = $pres.Slides.Add($pres.Slides.Count + 1, 12)
    
    # Add Title TextBox (Orientation: msoTextOrientationHorizontal = 1)
    $titleShape = $slide.Shapes.AddTextbox(1, 20, 10, $slideWidth - 40, 40)
    $titleShape.TextFrame.TextRange.Text = $item.Title
    $titleShape.TextFrame.TextRange.Font.Size = 24
    $titleShape.TextFrame.TextRange.Font.Bold = -1
    $titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # ppAlignCenter

    # Add Desc TextBox
    $descShape = $slide.Shapes.AddTextbox(1, 20, 50, $slideWidth - 40, 60)
    $descShape.TextFrame.TextRange.Text = $item.Desc
    $descShape.TextFrame.TextRange.Font.Size = 16
    $descShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # ppAlignCenter
    $descShape.TextFrame.WordWrap = -1

    # Add Image
    $imgPath = Join-Path -Path $baseDir -ChildPath "image\$($item.Image)"
    if (Test-Path $imgPath) {
        # AddPicture(FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
        # LinkToFile: msoFalse = 0, SaveWithDocument: msoTrue = -1
        $pic = $slide.Shapes.AddPicture($imgPath, 0, -1, 0, 0, -1, -1)
        
        $availWidth = $slideWidth - 40
        $availHeight = $slideHeight - 120
        
        $ratioW = $availWidth / $pic.Width
        $ratioH = $availHeight / $pic.Height
        $ratio = if ($ratioW -lt $ratioH) { $ratioW } else { $ratioH }
        
        $newWidth = $pic.Width * $ratio
        $newHeight = $pic.Height * $ratio
        
        $pic.LockAspectRatio = 0
        $pic.Width = $newWidth
        $pic.Height = $newHeight
        
        $pic.Left = ($slideWidth - $newWidth) / 2
        $pic.Top = 110 + ($availHeight - $newHeight) / 2
    }
}

$outputPath = Join-Path -Path $baseDir -ChildPath "Vimlesh_GitHub_Projects_Summary.pptx"
if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force
}

Write-Host "Saving presentation to $outputPath"
# 24 = ppSaveAsOpenXMLPresentation
$pres.SaveAs($outputPath, 24)
$pres.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host "Created PPTX successfully using PowerPoint COM object."
