Public Sub deleteOKs()
'
' Open 1 CSV file, load it, ,check if it is OK curve, close it, delete it and repete until list is empty or chart is full
'

Dim pathToCSVs As String
Dim toDelete As Boolean
Dim OKs As Integer, total As Integer

toDelete = False
pathToCSVs = ThisWorkbook.Path & "\CSV\"
OKs = 0
total = 0

CSVs = Dir(pathToCSVs & "*.csv")

Do While CSVs <> ""
    Workbooks.Open Filename:=pathToCSVs & CSVs, local:=True
    total = total + 1
    
    If Workbooks(CSVs).Worksheets(1).Range("B10").Value = "OK" Then
        toDelete = True
        OKs = OKs + 1
    End If
    
    ' Close the CSV file w/o saving changes
    Workbooks(CSVs).Close SaveChanges:=False
    
    ' Delete OK curve
    If toDelete = True Then
        Kill (pathToCSVs & CSVs)
        toDelete = False
    End If
    
    CSVs = Dir
Loop
MsgBox (total & " files were processed" & vbCrLf & _
    OKs & " files were deleted and " & total - OKs & " curves remained")
End Sub
