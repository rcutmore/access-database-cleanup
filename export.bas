Option Compare Database: Option Explicit

Public Sub ExportDatabaseObjects()
    On Error GoTo Err_ExportDatabaseObjects

    Dim db As DAO.Database
    Set db = CurrentDb()

    ' Make sure to include a closing back slash (ie: C:\Temp\).
    Dim exportLocation As String
    exportLocation = "Enter path to folder to store exported files"

    ' Export all tables.
    Dim td As TableDef
    Dim name As String
    For Each td In db.TableDefs
        Dim isValidTable As Boolean
        isValidTable = Left(td.name, 4) <> "MSys" And Left(td.name, 1) <> "~"

        If isValidTable Then
            DoCmd.TransferText _
                acExportDelim, , td.name, _
                exportLocation & "Table_" & td.name & ".txt", True
        End If
    Next td

    Set td = Nothing

    ' Export all forms.
    Dim c As Container
    Set c = db.Containers("Forms")
    Dim d As Document
    For Each d In c.Documents
        Application.SaveAsText _
            acForm, d.name, exportLocation & "Form_" & d.name & ".txt"
    Next d

    ' Export all reports.
    Set c = db.Containers("Reports")
    For Each d In c.Documents
        Application.SaveAsText _
            acReport, d.name, exportLocation & "Report_" & d.name & ".txt"
    Next d

    ' Export all macros.
    Set c = db.Containers("Scripts")
    For Each d In c.Documents
        Application.SaveAsText _
            acMacro, d.name, exportLocation & "Macro_" & d.name & ".txt"
    Next d

    ' Export all modules.
    Set c = db.Containers("Modules")
    For Each d In c.Documents
        Application.SaveAsText _
            acModule, d.name, exportLocation & "Module_" & d.name & ".txt"
    Next d

    Set c = Nothing
    Set d = Nothing

    ' Export all queries.
    Dim i As Integer
    For i = 0 To db.QueryDefs.Count - 1
        Application.SaveAsText _
            acQuery, db.QueryDefs(i).name, _
            exportLocation & "Query_" & db.QueryDefs(i).name & ".txt"
    Next i

    Set db = Nothing

    MsgBox "All database objects have been exported as a text file to " & _
           exportLocation, vbInformation, "Export complete"
    Exit Sub

Err_ExportDatabaseObjects:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Error"
End Sub
