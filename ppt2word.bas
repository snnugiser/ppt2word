Attribute VB_Name = "copy2word"
Option Explicit

Sub copy2word()
Dim pptPres As PowerPoint.Presentation
Set pptPres = Application.ActivePresentation

Dim word1 As Object
Set word1 = CreateObject("Word.Application")
word1.Visible = True
Dim doc1 As Variant
Set doc1 = word1.Documents.Add()
Dim slct1 As Variant
Set slct1 = word1.Selection
'slct1.typetext "hello word"

Dim temp As String

Dim pages As Integer
pages = pptPres.Slides.Count
Dim i As Integer
Dim j As Integer
Dim oslide As Slide
Dim oshape As Shape

For Each oslide In pptPres.Slides
        For Each oshape In oslide.Shapes
            If oshape.Type = msoTextBox Then
            'MsgBox pptPres.Slides(i).Shapes(j).TextFrame.TextRange.Text
                slct1.InsertAfter oshape.TextFrame.TextRange.Text
                slct1.InsertAfter vbCrLf
            End If

        Next
Next
doc1.SaveAs ("C:\doc.doc")
doc1.Close
word1.Quit
End Sub
