Attribute VB_Name = "Module1"

Public con As New ADODB.Connection
Public Sub connect()
    con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\LMS.accdb;"
    con.CursorLocation = adUseClient
    con.Open
    
End Sub


