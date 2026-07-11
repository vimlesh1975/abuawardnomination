const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const data = [
  { title: "DeckLink Vertical Scroll Player", desc: "UI for a vertical scroll player with DeckLink device selection, text formatting options, and scrolling text about sports matches.", image: "decklink_vertical_scroll_player.png" },
  { title: "News Scroll Scheduler", desc: "UI for scheduling news scrolls with tabs for various templates, displaying multilingual text in a grid.", image: "casparcg_news_scroll.png" },
  { title: "DeckLink Audio Recorder", desc: "UI for audio recording, showing hardware configuration and live audio level meters for 14 channels.", image: "decklink_audio_recorder.png" },
  { title: "CasparCG Media Playback", desc: "CasparCG media playback interface with playlist of .mp4 videos and server connection status.", image: "caspar_media_playback.png" },
  { title: "DecklinkPlayer Manual Mode", desc: "UI showing DecklinkPlayer in manual mode, playlist of MXF video files, and playback controls.", image: "decklink_player.png" },
  { title: "Web Teleprompter Interface", desc: "Web-based teleprompter interface displaying Hindi text with speed controls and reading options.", image: "web_teleprompter.png" },
  { title: "Samvad WebSocket Inspector", desc: "Web UI for managing stories and run orders for a teleprompter, showing Hindi text content.", image: "samvad_websocket_inspector.png" },
  { title: "4 Channel DeckLink Recorder", desc: "UI for a 4-channel video recording application, showing 4 simultaneous video inputs from DeckLink Duo cards.", image: "4_channel_decklink_recorder.png" },
  { title: "CefToDecklink HTML Renderer", desc: "UI for HTML to DeckLink Renderer, showing HTML URL pointing to CasparcgOutput and a live preview.", image: "cef_to_decklink.png" },
  { title: "CasparCG Scroll Scheduler", desc: "Web UI for scheduling scrolling text on CasparCG, showing fields for host, font family, layer, and speed.", image: "casparcg_scroll_scheduler.png" },
  { title: "Streaming Studio Dashboard", desc: "Web-based studio interface showing live guest video feeds and invite link generator.", image: "webrtc_studio_getstreamio.png" },
  { title: "Audio Off-Air Logger", desc: "UI showing audio input device, waveform, and a list of saved audio clips with download options.", image: "audio_off_air_logger.png" },
  { title: "ReactCasparClient", desc: "A web-based client for CasparCG showing designer tools and a live output preview.", image: "react_caspar_client_designer.png" },
  { title: "WebAnimator for CasparCG", desc: "Web-based animation tool for CasparCG graphics showing a timeline, keyframes, and a preview of a sports graphic.", image: "react_caspar_client_webanimator.png" },
  { title: "CasparCG Election Control Panel", desc: "Web dashboard for controlling election graphics, showing party data and seats controls.", image: "casparcg_election_control_panel.png" },
  { title: "Media Explorer and Screen Consumer", desc: "UI showing a media explorer file tree and a grid interface for playing video blocks.", image: "project_screenshot_16.png" },
  { title: "Playout Counter", desc: "Application showing countdown timers for News, Commercial, and Promo playout segments.", image: "project_screenshot_17.png" },
  { title: "DD Chennai Quiz Graphics", desc: "Web dashboard for triggering quiz graphics showing a gallery of images and video files.", image: "project_screenshot_18.png" },
  { title: "Web Audio Recorder", desc: "Simple web-based audio recorder showing a live waveform and a list of past recordings.", image: "project_screenshot_19.png" },
  { title: "Election Data Spreadsheet Dashboard", desc: "UI showing tables with municipal election data, listing parties and vote counts for various regions.", image: "project_screenshot_20.png" },
  { title: "Election Results Graphic Preview", desc: "A full-screen graphic displaying Marathi text and tables with election results for various regions.", image: "project_screenshot_21.png" }
];

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // default is 16x9: 10 x 5.625 inches
pres.company = 'Vimlesh';
pres.title = 'GitHub Projects Summary';

data.forEach((item, index) => {
    let slide = pres.addSlide();
    
    // Title
    slide.addText(item.title, { 
        x: 0.5, y: 0.2, w: 9.0, h: 0.5, 
        fontSize: 24, bold: true, color: '363636', align: 'center' 
    });
    
    // Description
    slide.addText(item.desc, { 
        x: 0.5, y: 0.8, w: 9.0, h: 0.8, 
        fontSize: 16, color: '666666', align: 'center', wrap: true 
    });
    
    // Image
    const imgPath = path.join(__dirname, 'image', item.image);
    if (fs.existsSync(imgPath)) {
        slide.addImage({ 
            path: imgPath, 
            x: 0.5, y: 1.7, w: 9.0, h: 3.5, 
            sizing: { type: 'contain', w: 9.0, h: 3.5 } 
        });
    } else {
        console.warn("Image not found: " + imgPath);
    }
});

const outputPath = path.join(__dirname, 'Vimlesh_GitHub_Projects_Summary.pptx');

pres.writeFile({ fileName: outputPath }).then(fileName => {
    console.log(`created file: ${fileName}`);
}).catch(err => {
    console.error("Error creating PPTX: ", err);
});
