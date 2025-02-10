# Pricing on monetary notes using VBA and SQL

This project was divided into three main parts:

1) Data Retrieval: First, we retrieved data from the Bloomberg interface. We created an Excel file named "MONETARY_INDEX_DATA", containing the values of multiple monetary indices such as EONIA and ESTER from 01/01/2016 to 12/31/2023. We then compiled another file, "NOTES_DATA", which includes the key characteristics of the issued notes. Finally, we generated a third file, "NOTE_RATE_DATA", consolidating all relevant rate information (using the files "spread" and "rerating").
   
2) Database construction: Using the collected data, we built a database in SQL with four distinct tables:

- "MONETARY_INDEX"
- "ISSUANCE_STATIC_DATA"
- "ISSUANCE_RATE_DATA"
- "ISSUANCE_DYNAMIC_DATA"

The "ISSUANCE_DYNAMIC_DATA" table was populated using data from the other three tables.
To construct these tables efficiently, we leveraged VBA. Below is the VBA code used to create the first table, which served as a template for the others:

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
        TableSQL = "INSERT INTO " & table_name & " VALUES ('" & rgData.Cells(i, 1).Value & "', '" & rgData.Cells(i, 2).Value & "', '" & rgData.Cells(i, 3).Value & "', '" &                                 rgData.Cells(i, 4).Value & "', '" & rgData.Cells(i, 5).Value & "', '" & rgData.Cells(i, 6).Value & "', '" & rgData.Cells(i, 7).Value & "')"
        
        ' Execution of the request
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

To execute the SQL request in VBA we made this public function:
    
    Public Function RunSqlRequest(sRequest As String, sPathDB As String) As ADODB.Recordset
     
    Dim conn As New ADODB.Connection
    Dim rec As New ADODB.Recordset
     
    conn.Open "Provider=Microsoft.ACE.OLEDB.16.0;data source=" & sPathDB
        
    Set rec = conn.Execute(sRequest)
    Set RunSqlRequest = rec
  
    Set conn = Nothing
       
    End Function

3) Lastly, we created a file named "NOTE_TEMPLATE", containing a single sheet called "Note". On this sheet, entering an ISIN and a pricing date automatically retrieves and displays all relevant characteristics—such as the issue date, price, or potential rerating date—directly from the database.

To enhance the project's efficiency, we integrated four VBA buttons:
- One to clear all fields after a search is completed.
- One to print the document.
- One to send an email.
- One to generate a graph using the retrieved data.
