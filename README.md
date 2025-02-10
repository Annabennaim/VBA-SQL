# Pricing on monetary notes using VBA and SQL

This project was decomposed in three main parts:

1) First, we had to retrieve data from the interface Bloomberg. We created an excel file named "MONETARY_INDEX_DATA" with the values of multiple monetay index such as EONIA or ESTER, from 01/01/2016 to 12/31/2023. We used the data from Bloomberg. We then created an other file called "NOTES_DATA", which contains the principal caracteristics of the notes given. Finally, we created a final file named " NOTE_RATE_DATA".
2) Secondly, with alle the date we retrieve, we had to construct a database on SQL with 4 different tables: "MONETARY_INDEX", "ISSUANCE_STATIC_DATA", "ISSUANCE_RATE_DATA" AND "ISSUANCE_DYNAMIC_DATA". This last table was fill with the data of the other tables.
To construct these tables, we used VBA and here is the code (this is for the first table but we used the same "main" code for the others):
Public Sub CreerTable MONETARY_INDEX_DATA()

'Here are the variables
    Dim pathdonnees As String, pathaccess As String
    Dim wbdata As Workbook
    Dim Monetary_Index As Worksheet
    Dim TableSQL As String, InsertSQL As String
    Dim table_name As String
    Dim rgFields As Range
    Dim rgData As Range
    Dim i As Long, j As Long
    Dim rec As New ADODB.Recordset
    Dim vbdata As Date, vbdouble As Double
    
    Application.ScreenUpdating = False

    ' To open the file with the data (Monetary Index)
    pathdonnees = Application.GetOpenFilename
    Set wbdata = Workbooks.Open(pathdonnees)
    
    ' Definition of the sheet where we want to put the data
    Set Monetary_Index = wbdata.Worksheets(1)
    
    Set rgData = Monetary_Index.Range("A7")
    Set rgData = Range(rgData, rgData.End(xlDown).End(xlDown).End(xlToRight))

    ' Access database path
    pathaccess = Application.GetOpenFilename

    ' Name of the table
    table_name = ThisWorkbook.Worksheets(1).Range("TABLE1").Value

    ' Definition of the fields of the table
    Set rgFields = ThisWorkbook.Worksheets(1).Range("TABLE1")
    Set rgFields = rgFields.Offset(1, -1)
    Set rgFields = Range(rgFields, rgFields.End(xlDown).End(xlToRight))

    TableSQL = "CREATE TABLE [" & table_name & "] ("
    For i = 1 To rgFields.Rows.Count
        TableSQL = TableSQL & "[" & rgFields(i, 1).Value & "] "
        TableSQL = TableSQL & rgFields(i, 2).Value & " NOT NULL, "
    Next i
    
    TableSQL = Left(TableSQL, Len(TableSQL) - 2) & ")"

    ' Execution of the request
    On Error GoTo 0
    Set rec = RunSqlRequest(TableSQL, pathaccess)

    ' Résults
    MsgBox "Table créée avec succés.", vbInformationn
    
' Boucle pour remplir la table
    For i = 1 To rgData.Rows.Count - 6
        ' Construction de la requte en ins_rant les diff_rents champs de la tablee
        TableSQL = "INSERT INTO " & table_name & " VALUES ('" & rgData.Cells(i, 1).Value & "', '" & rgData.Cells(i, 2).Value & "', '" & rgData.Cells(i, 3).Value & "', '" & rgData.Cells(i, 4).Value & "', '" & rgData.Cells(i, 5).Value & "', '" & rgData.Cells(i, 6).Value & "', '" & rgData.Cells(i, 7).Value & "')"
        
        ' Execution de la requete
        On Error Resume Next
        RunSqlRequest TableSQL, pathaccess
        If Err.Number <> 0 Then
            MsgBox "Erreur lors de l'insertion des données ˆ la ligne " & i & ": " & Err.Description, vbCritical
            Err.Clear
            On Error GoTo 0
            Exit For ' Stop si erreur
        End If
        On Error GoTo 0 'on reparamtre la d_tection d'erreurss
    Next i

    ' Affichage pour débogage
    Debug.Print TableSQL

    ' Dégel de l'_cran
    Application.ScreenUpdating = True
End Sub




4) Lastly, we had to create an other file with just one sheet "Note"
