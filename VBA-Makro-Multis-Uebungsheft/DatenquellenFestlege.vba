Option Explicit

Sub Merge_To_Individual_Files()
'Merges one record at a time to the chosen output folder

Application.ScreenUpdating = False
Dim StrFolder As String, StrName As String, MainDoc As Document, i As Long, j As Long, a As String
Set MainDoc = ActiveDocument
a = MsgBox("Möchten Sie die Übungsdateien für die Teilnehmer erstellen?", vbYesNo)
If a = vbNo Then
    a = MsgBox("Möchten Sie die Übungen drucken?", vbYesNo)
    If a = vbNo Then Exit Sub
    Call DatenquelleFestlegen
    Call SerieDrucken(MainDoc)
    ActiveDocument.MailMerge.MainDocumentType = wdNotAMergeDocument
    If Application.Documents.Count = 1 Then Application.Quit SaveChanges:=False
    ActiveDocument.Close SaveChanges:=False
    Exit Sub
End If
With ActiveWindow.View.RevisionsFilter 'Kommentare für den Druck ausschalten
        .Markup = wdRevisionsMarkupNone
        .View = wdRevisionsViewFinal
End With
Call DatenquelleFestlegen
With MainDoc
  For i = 1 To .MailMerge.DataSource.RecordCount
    With .MailMerge
      .Destination = wdSendToNewDocument
      .SuppressBlankLines = True
      With .DataSource
        .FirstRecord = i
        .LastRecord = i
        .ActiveRecord = i
        StrFolder = MainDoc.PATH & "\Teilnehmerordner\Teilnehmer " & .DataFields("ID") & Application.PathSeparator
        StrName = Left(ThisDocument.Name, InStrRev(ThisDocument.Name, "-")) & " TN" & .DataFields("ID")
        
      End With
      .Execute Pause:=False
    End With
    StrName = Trim(StrName)
    With ActiveDocument
      .SaveAs2 FileName:=StrFolder & StrName & ".pdf", FileFormat:=wdFormatPDF, AddToRecentFiles:=False
      .Close SaveChanges:=False
    End With
  Next i
  
End With
Application.ScreenUpdating = True
MsgBox "Es wurden " & i - 1 & " Übungsdateien erstellt und in den entsprechenden Teilnehmerordnern gespeichert."
a = MsgBox("Möchten Sie die Übungen drucken?", vbYesNo)
If a = vbYes Then Call SerieDrucken(MainDoc)
ActiveDocument.MailMerge.MainDocumentType = wdNotAMergeDocument
If Application.Documents.Count = 1 Then Application.Quit SaveChanges:=False
ActiveDocument.Close SaveChanges:=False

End Sub
Sub DatenquelleFestlegen()
Dim PATH As String, SOURCE As String
PATH = ThisDocument.PATH 'determine current directory
SOURCE = PATH & "\Übungsdaten.xlsx"
ActiveDocument.MailMerge.OpenDataSource Name:=SOURCE, SQLStatement:="SELECT * FROM `Grundmodul_Übungsheft$`"
End Sub

Sub SerieDrucken(Dokument As Document)
    SendKeys "%+m", True
    DoEvents
    Do Until Application.ActiveWindow.Caption = Dokument.Name
     DoEvents
    Loop
End Sub

Private Sub Document_Open()
    Call Merge_To_Individual_Files
End Sub










