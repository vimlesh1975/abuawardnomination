Set objPPT = CreateObject("PowerPoint.Application")
objPPT.Visible = True

Dim pptxPath, pdfPath
pptxPath = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination\Vimlesh_GitHub_Projects_Summary.pptx"
pdfPath = "C:\Users\vimlesh\Documents\vimlesh\abuawardnomination\Vimlesh_GitHub_Projects_Summary.pdf"

' Open the existing presentation
Set objPres = objPPT.Presentations.Open(pptxPath, True, False, False)

On Error Resume Next

' Try SaveAs first (format 32 = ppSaveAsPDF)
objPres.SaveAs pdfPath, 32

If Err.Number <> 0 Then
    Err.Clear
    ' Try ExportAsFixedFormat (format 2 = ppFixedFormatTypePDF, 1 = ppFixedFormatIntentScreen)
    objPres.ExportAsFixedFormat pdfPath, 2, 1
End If

objPres.Close
objPPT.Quit

WScript.Echo "Created PDF successfully."
