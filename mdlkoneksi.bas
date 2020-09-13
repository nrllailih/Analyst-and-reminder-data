Attribute VB_Name = "Module1"
Option Explicit
Public koneksi As New ADODB.Connection
Public Rsaskrindo As New ADODB.Recordset
Public RsAdmin As New ADODB.Recordset
Public rshasil As New ADODB.Recordset
Public Rs, Rs_Data As New ADODB.Recordset
Public Sql, strSQL As String
Public uang1, uang2, uang3, uang4 As Currency
Public tgl1, tgl2, tgl3, tgl4 As Date

Public Sub konekdb()
Set koneksi = New ADODB.Connection
koneksi.CursorLocation = adUseClient

koneksi = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbpkl.mdb; Mode=readwrite"

If koneksi.State = adStateOpen Then
    koneksi.Close
    Set koneksi = New ADODB.Connection
    koneksi.Open koneksi
Else
    koneksi.Open koneksi
End If
If Err.Number <> 0 Then
    MsgBox "Koneksi database gagal, periksa modul koneksi", vbCritical, "Error"
    Exit Sub
Else
    MsgBox "Koneksi Berhasil", vbInformation, "Sukses"
End If



End Sub







