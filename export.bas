Public Sub ExportDatabaseObjects()
On Error GoTo Err_ExportDatabaseObjects

    Dim db As DAO.Database
    Dim td As TableDef
    Dim d As Document
    Dim c As Container
    Dim i As Integer
    Dim exportLocation As String
    Dim name As String

    Set db = CurrentDb()

    ' Make sure to include a closing back slash. (ie: C:\Temp\)
    exportLocation = "Enter path to folder to store exported files"

    For Each td In db.TableDefs
        If Left(td.name, 4) <> "MSys" And Left(td.name, 1) <> "~" Then
            DoCmd.TransferText acExportDelim, , td.name, exportLocation & "Table_" & td.name & ".txt", True
        End If
    Next td

    Set c = db.Containers("Forms")
    For Each d In c.Documents
        Application.SaveAsText acForm, d.name, exportLocation & "Form_" & d.name & ".txt"
    Next d

    Set c = db.Containers("Reports")
    For Each d In c.Documents
        Application.SaveAsText acReport, d.name, exportLocation & "Report_" & d.name & ".txt"
    Next d

    Set c = db.Containers("Scripts")
    For Each d In c.Documents
        Application.SaveAsText acMacro, d.name, exportLocation & "Macro_" & d.name & ".txt"
    Next d

    Set c = db.Containers("Modules")
    For Each d In c.Documents
        Application.SaveAsText acModule, d.name, exportLocation & "Module_" & d.name & ".txt"
    Next d

    For i = 0 To db.QueryDefs.count - 1
        Application.SaveAsText acQuery, db.QueryDefs(i).name, exportLocation & "Query_" & db.QueryDefs(i).name & ".txt"
    Next i

    Set db = Nothing
    Set c = Nothing

    MsgBox "All database objects have been exported as a text file to " & exportLocation, vbInformation

Exit_ExportDatabaseObjects:
    Exit Sub

Err_ExportDatabaseObjects:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ExportDatabaseObjects
End Sub