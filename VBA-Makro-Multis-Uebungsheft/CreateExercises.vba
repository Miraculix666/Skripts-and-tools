Option Explicit

Sub Merge_To_Individual_Files()
    ' Merges one record at a time to the chosen output folder

    Application.ScreenUpdating = False
    Dim strFolder As String, strName As String, MainDoc As Document, i As Long, a As String
    Dim participantList As Object ' New object to hold participant list
    Set MainDoc = ActiveDocument
    
    a = MsgBox("Do you want to create exercise files for the participants?", vbYesNo)
    If a = vbNo Then
        a = MsgBox("Do you want to print the exercises?", vbYesNo)
        If a = vbNo Then Exit Sub
        Call SetDataSource
        Call PrintSeries(MainDoc)
        ActiveDocument.MailMerge.MainDocumentType = wdNotAMergeDocument
        If Application.Documents.Count = 1 Then Application.Quit SaveChanges:=False
        ActiveDocument.Close SaveChanges:=False
        Exit Sub
    End If
    
    With ActiveWindow.View.RevisionsFilter ' Turn off comments for printing
        .Markup = wdRevisionsMarkupNone
        .View = wdRevisionsViewFinal
    End With
    
    Call SetDataSource
    Call LoadParticipantList ' Placeholder for participant list loading logic
    
    With MainDoc
        For i = 1 To .MailMerge.DataSource.RecordCount
            With .MailMerge
                .Destination = wdSendToNewDocument
                .SuppressBlankLines = True
                
                With .DataSource
                    .FirstRecord = i
                    .LastRecord = i
                    .ActiveRecord = i
                    
                    strFolder = MainDoc.Path & "\ParticipantFolders\Participant_" & .DataFields("ID") & Application.PathSeparator
                    strName = Left(ThisDocument.Name, InStrRev(ThisDocument.Name, "-")) & " P" & .DataFields("ID")
                    
                    ' Create folder if it doesn't exist
                    If Dir(strFolder, vbDirectory) = "" Then MkDir strFolder
                End With
                
                .Execute Pause:=False
            End With
            
            strName = Trim(strName)
            With ActiveDocument
                .SaveAs2 FileName:=strFolder & strName & ".pdf", FileFormat:=wdFormatPDF, AddToRecentFiles:=False
                .Close SaveChanges:=False
            End With
        Next i
        
    End With
    
    Application.ScreenUpdating = True
    MsgBox "Created " & i - 1 & " exercise files and saved them in the corresponding participant folders."
    
    a = MsgBox("Do you want to print the exercises?", vbYesNo)
    If a = vbYes Then Call PrintSeries(MainDoc)
    
    ActiveDocument.MailMerge.MainDocumentType = wdNotAMergeDocument
    
    If Application.Documents.Count = 1 Then Application.Quit SaveChanges:=False
    ActiveDocument.Close SaveChanges:=False

End Sub

Sub SetDataSource()
    Dim path As String, source As String
    Dim fd As FileDialog
    
    path = ThisDocument.Path ' Determine current directory
    source = path & "\ExerciseData.xlsx"
    
    ' Check if the file exists
    If Dir(source) = "" Then
        ' File not found, open file dialog to select data source manually
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
            .Title = "Please select the data source"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xlsx; *.xls", 1
            .InitialFileName = path
            
            If .Show = -1 Then ' User selected a file
                source = .SelectedItems(1)
            Else ' User cancelled the dialog box
                MsgBox "No data source selected. The process will be aborted.", vbExclamation
                Exit Sub
            End If
        End With
        
    End If
    
    ' Try to open the data source
    On Error Resume Next
    ActiveDocument.MailMerge.OpenDataSource Name:=source, SQLStatement:="SELECT * FROM `BasicModule_ExerciseBook$`"
    
    If Err.Number <> 0 Then ' Handle errors when opening the data source fails.
        MsgBox "Error opening data source: " & Err.Description, vbCritical
        Err.Clear: Exit Sub
    End If
    
    On Error GoTo 0
    
End Sub

Sub PrintSeries(Dokument As Document)
    SendKeys "%+m", True ' Simulate Alt+Shift+M for Mail Merge Print Preview shortcut (if applicable)
    DoEvents
    
    Do Until Application.ActiveWindow.Caption = Dokument.Name ' Wait until focus returns to document window.
        DoEvents
    Loop

End Sub

Private Sub Document_Open()
    Call Merge_To_Individual_Files ' Automatically execute main routine when document opens.
End Sub

Sub LoadParticipantList()
    ' Placeholder for loading participant list from an additional data source.
End Sub

Sub GenerateConsistentDataset()
    ' Placeholder for generating consistent datasets across documents.
End Sub

Sub MoodleIntegration()
    ' Placeholder for integrating with Moodle using participant IDs.
End Sub

