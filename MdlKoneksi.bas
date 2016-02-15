Attribute VB_Name = "MdlKoneksi"
Public hubungkan As String
Public RSADMIN As New ADODB.Recordset

Public Sub koneksi()
Set RSADMIN = New ADODB.Recordset
hubungkan = "provider=microsoft.jet.oledb.4.0 ; data source = " & App.Path & "/database.mdb;"
End Sub

