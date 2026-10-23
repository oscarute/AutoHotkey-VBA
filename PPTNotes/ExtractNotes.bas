Attribute VB_Name = "ExtractNotes"
' Extract all speaker notes of the active presentation into a .txt file
' saved next to the presentation, then open it in Notepad.
' The presentation must be saved before running this macro.

Sub ExtractNotes()
    Dim s As Slide, shp As Shape
    Dim txt As String, notes As String, outPath As String, f As Integer

    For Each s In ActivePresentation.Slides
        notes = ""
        For Each shp In s.NotesPage.Shapes
            ' Find the notes box by type, not by position (Placeholders(2) fails on some slides)
            If shp.Type = msoPlaceholder Then
                If shp.PlaceholderFormat.Type = ppPlaceholderBody Then
                    If shp.HasTextFrame Then notes = shp.TextFrame.TextRange.Text
                End If
            End If
        Next shp
        txt = txt & "=== Slide " & s.SlideIndex & " ===" & vbCrLf & _
              IIf(notes = "", "(no notes)", notes) & vbCrLf & vbCrLf
    Next s

    outPath = ActivePresentation.Path & "\Notes_" & _
              Left(ActivePresentation.Name, InStrRev(ActivePresentation.Name, ".") - 1) & ".txt"
    f = FreeFile
    Open outPath For Output As #f
    Print #f, txt
    Close #f

    Shell "notepad.exe """ & outPath & """", vbNormalFocus
End Sub
