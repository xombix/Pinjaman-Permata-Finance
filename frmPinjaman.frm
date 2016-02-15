VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPinjaman 
   Caption         =   "Form1"
   ClientHeight    =   8955
   ClientLeft      =   5100
   ClientTop       =   1230
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   10485
   Begin VB.Frame Frame1 
      Caption         =   "Form1"
      Height          =   7935
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   8520
         Top             =   480
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1560
         TabIndex        =   27
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   735
         Left            =   1560
         TabIndex        =   24
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Left            =   7560
         Top             =   240
      End
      Begin VB.TextBox txtTgl 
         Height          =   375
         Left            =   5280
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   6360
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Tambah"
         Height          =   495
         Left            =   4080
         TabIndex        =   21
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Batal"
         Height          =   495
         Left            =   4080
         TabIndex        =   20
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   5640
         TabIndex        =   19
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   5640
         TabIndex        =   18
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Simpan"
         Height          =   495
         Left            =   7440
         TabIndex        =   17
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Keluar"
         Height          =   495
         Left            =   7440
         TabIndex        =   16
         Top             =   2640
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   6360
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   7800
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Hitung"
         Height          =   495
         Left            =   4080
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmPinjaman.frx":0000
         Left            =   1560
         List            =   "frmPinjaman.frx":0002
         TabIndex        =   6
         Top             =   2640
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   7680
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   1560
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid gridPinjaman 
         Bindings        =   "frmPinjaman.frx":0004
         Height          =   2655
         Left            =   120
         TabIndex        =   15
         Top             =   5040
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4683
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label11 
         Caption         =   "Surveyor"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Note"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Tanggal"
         Height          =   375
         Left            =   4200
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Total Pinjaman"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   4560
         Width           =   6615
      End
      Begin VB.Label Label7 
         Caption         =   "ID Pinjaman"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Nasabah"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Cicilan Perbulan : "
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "Pinjaman"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Tenor"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Tahun Kendaraan"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Jenis Kendaraan"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmPinjaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sHari As String
Dim aHari
Private Sub idPinjam()
'id Pinjaman
Dim intCount As Integer
Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc4.RecordSource = "select*from pinjaman"
Adodc4.Refresh
'intCount = Adodc4.RecordCount
For i = 1 To 5
X = Adodc4.Recordset!id_pinjaman
If X <> i Then
NewEmpID = X ' 'i' can also be used
Exit For
End If
Next i

FinalEmpID = Format(NewEmpID + 1, "00000")
Text1.Text = FinalEmpID
'Now you just have to add the FinalEmpID as the new employee ID

End Sub
Private Sub kosong()
Text2.Text = ""
Text2.SetFocus
End Sub

Private Sub cmdAdd_Click()
kosong
cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdExit.Enabled = False
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.SetFocus
    Adodc4.Recordset.AddNew
    gridPinjaman.AllowUpdate = True

End Sub

Private Sub cmdCancel_Click()
kosong
cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdDelete.Enabled = True
    cmdExit.Enabled = True
    Adodc4.Recordset.Cancel
    Adodc4.Refresh
    gridPinjaman.AllowUpdate = False
End Sub

Private Sub cmdDelete_Click()
 Dim result As Integer
    result = MsgBox("Hapus data ini?", vbOKCancel, "Konfirmasi")
    If result = 2 Then
        Adodc4.Recordset.CancelUpdate
            Else
        If Not Adodc4.Recordset.EOF Then
            Adodc4.Recordset.Delete
            Adodc4.Recordset.MoveFirst
        End If
    End If

End Sub

Private Sub cmdEdit_Click()
cmdEdit.Enabled = False
    cmdSave.Enabled = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdExit.Enabled = False
    cmdSave.SetFocus
    gridPinjaman.AllowUpdate = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
 
 On Error GoTo pesan
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    Dim result As Integer
    result = MsgBox("Perbaharui data ini?", vbOKCancel, "Konfirmasi")
    If result = 2 Then
        Call cmdCancel_Click
    Else
    If Not Text2.Text = "" Then
total = Combo2.ItemData(Combo2.ListIndex) + Combo1.ItemData(Combo1.ListIndex)
Gtotal = ((Val(total) / Val(Combo3.ItemData(Combo3.ListIndex))) + (0.05 * Val(total))) / Val(Combo3.ItemData(Combo3.ListIndex))
        Adodc4.Recordset!id_pinjaman = Text1.Text
        Adodc4.Recordset!nama = Combo4.Text
        Adodc4.Recordset!plafon = Combo3.ItemData(Combo3.ListIndex)
        Adodc4.Recordset!surveyor = Combo5.Text
        Adodc4.Recordset!tPinjaman = Gtotal
       Adodc4.Recordset!note = Text2.Text
       Adodc4.Recordset!tgl = txtTgl.Text
       Adodc4.Recordset!jenis = Combo1.ItemData(Combo1.ListIndex)
       Adodc4.Recordset!tahun = Combo2.ItemData(Combo2.ListIndex)
        Adodc4.Recordset.Update
        cmdSave.Enabled = False
        cmdDelete.Enabled = True
        cmdExit.Enabled = True
        gridPinjaman.AllowUpdate = False
    Else
        result = MsgBox("Data Harus Lengkap", vbInformation, "Informasi")
        cmdAdd.Enabled = False
    End If
    End If
pesan:
    Select Case Err.Number
        Case -2147467259
        MsgBox "Data Sudah Ada!", vbCritical, "Perhatian"
        Call cmdCancel_Click
    End Select
idPinjam
End Sub






Private Sub gridPinjaman_Click()
Text1.Text = gridPinjaman.Columns(0).Text
Text2.Text = gridPinjaman.Columns(5).Text
End Sub



Private Sub Command1_Click()
'Label4.Caption = Combo3.List(Combo3.ListIndex)
'If Not Val(Combo1.ItemData(Combo1.ListIndex)) = Null And Not Val(Combo2.ItemData(Combo2.ListIndex)) = Null And Not Val(Combo3.ItemData(Combo3.ListIndex)) = Null Then
Label4.Caption = "Pinjaman :  Rp. " & Format(Val(Combo2.ItemData(Combo2.ListIndex)) + Val(Combo1.ItemData(Combo1.ListIndex)), "###,###,##0.00")
total = Combo2.ItemData(Combo2.ListIndex) + Combo1.ItemData(Combo1.ListIndex)
Gtotal = (Val(total) / Val(Combo3.ItemData(Combo3.ListIndex))) + (0.05 * Val(total))
Label5.Caption = "Jumlah Angsuran Perbulan :  Rp. " & Format(Round(Gtotal), "###,###,##0.00")
Label8.Caption = "Total Pinjaman :  Rp. " & Format(Round(Gtotal * Val(Combo3.ItemData(Combo3.ListIndex))), "###,###,##0.00")
'Else
 '       result = MsgBox("Data Harus Lengkap", vbInformation, "Informasi")
  '      Command1.Enabled = False
   ' End If
End Sub
Private Sub Timer1_Timer()
sHari = aHari(Abs(Weekday(Date) - 1))
txtTgl.Text = Format(Date, "m/d/yyyy")
End Sub

Private Sub Form_Load()
'load tanggal
idPinjam
aHari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True



'Combobox u/ jenis kendaraan
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc1.RecordSource = "select*from jenis_kendaraan "
Adodc1.Refresh
Combo1.Clear
Do While Not Adodc1.Recordset.EOF = True
        Combo1.AddItem Adodc1.Recordset!jenis
        Combo1.ItemData(Combo1.NewIndex) = Adodc1.Recordset!harga
        Adodc1.Recordset.MoveNext
 Loop
'Combobox u/ tahun kendaraan
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc2.RecordSource = "select*from thn_kendaraan "
Adodc2.Refresh
Combo2.Clear
Do While Not Adodc2.Recordset.EOF = True
        Combo2.AddItem Adodc2.Recordset!tahun
        Combo2.ItemData(Combo2.NewIndex) = Adodc2.Recordset!harga
        Adodc2.Recordset.MoveNext
 Loop
 'combobox tenor
Combo3.Clear
Combo3.AddItem "3 bulan"
Combo3.ItemData(Combo3.NewIndex) = 3
Combo3.AddItem "10 bulan"
Combo3.ItemData(Combo3.NewIndex) = 10
Combo3.AddItem "18 bulan"
Combo3.ItemData(Combo3.NewIndex) = 18

'Combobox u/ nasabah
Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc3.RecordSource = "select*from nasabah "
Adodc3.Refresh
Combo4.Clear
Do While Not Adodc3.Recordset.EOF = True
        Combo4.AddItem Adodc3.Recordset!nama
       ' Combo4.ItemData(Combo4.NewIndex) = Adodc3.Recordset!nama
        Adodc3.Recordset.MoveNext
 Loop
 'Combobox u/ jenis kendaraan
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc1.RecordSource = "select*from jenis_kendaraan "
Adodc1.Refresh
Combo1.Clear
Do While Not Adodc1.Recordset.EOF = True
        Combo1.AddItem Adodc1.Recordset!jenis
        Combo1.ItemData(Combo1.NewIndex) = Adodc1.Recordset!harga
        Adodc1.Recordset.MoveNext
 Loop
'Combobox u/ surveyor
Adodc5.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc5.RecordSource = "select*from surveyor "
Adodc5.Refresh
Combo5.Clear
Do While Not Adodc5.Recordset.EOF = True
        Combo5.AddItem Adodc5.Recordset!nama
       Combo5.ItemData(Combo5.NewIndex) = Val(Adodc5.Recordset!nama)
        Adodc5.Recordset.MoveNext
 Loop
End Sub

