VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCicilan 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Form1"
      Height          =   7935
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bayar"
         Height          =   495
         Left            =   4080
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Keluar"
         Height          =   495
         Left            =   7440
         TabIndex        =   7
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Simpan"
         Height          =   495
         Left            =   7440
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   5640
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   5640
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Batal"
         Height          =   495
         Left            =   4080
         TabIndex        =   3
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Tambah"
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtTgl 
         Height          =   375
         Left            =   5280
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Timer Timer1 
         Left            =   7560
         Top             =   240
      End
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
      Begin MSDataGridLib.DataGrid gridCicilan 
         Bindings        =   "frmCicilan.frx":0000
         Height          =   2655
         Left            =   120
         TabIndex        =   10
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
      Begin VB.Label Label4 
         Caption         =   "Cicilan Ke-"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Total Bayar"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Denda"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Cicilan"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "ID Pinjaman"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Tanggal"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Nama nasabah"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCicilan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sHari As String
Dim aHari
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
    gridCicilan.AllowUpdate = True

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
    gridCicilan.AllowUpdate = False
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
    gridCicilan.AllowUpdate = True
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
        gridCicilan.AllowUpdate = False
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
End Sub
Private Sub Combo1_LostFocus()
Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc2.RecordSource = "select*from pinjaman where id_pinjaman = '" & Combo1.Text & "'"
Adodc2.Refresh
Text1.Text = Adodc2.Recordset!nama
Text2.Text = Format(Round(Adodc2.Recordset!tPinjaman), "###,###,##0.00")
date1 = Adodc2.Recordset!tgl
date2 = txtTgl.Text
'hitung lebih dari 2 hari
If DateDiff("d", date1, date2) > 2 Then
       hariH = DateDiff("d", date1, date2)
     Tdenda = Adodc2.Recordset!tPinjaman + ((Adodc2.Recordset!tPinjaman * 0.5) / 100)
     denda = ((Adodc2.Recordset!tPinjaman * 0.5) / 100)
     '  MsgBox ("Not Vending")
    End If
Text3.Text = Format(Round(denda), "###,###,##0.00")
Text4.Text = Format(Round(Adodc2.Recordset!tPinjaman + denda), "###,###,##0.00")
'hitung cicilan ke-
Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc3.RecordSource = "select count(id_pinjaman) as iPinj from cicilan where id_pinjaman = '" & Combo1.Text & "'"
Adodc3.Refresh
If Adodc3.Recordset!iPinj < 1 Then
CicilanK = 1
Else
CicilanK = Val(Adodc3.Recordset!iPinj) + 1
End If
Text5.Text = CicilanK


End Sub


Private Sub gridCicilan_Click()
Combo1.Text = gridCicilan.Columns(0).Text
Text5.Text = gridCicilan.Columns(1).Text
End Sub



Private Sub Command1_Click()
   'periksa jika jumlah cicilan sudah cukup
    'If Adodc2.Recordset!plafon > Text5.Text Then
     
     Adodc4.Recordset.AddNew
    gridCicilan.AllowUpdate = True
 On Error GoTo pesan
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    Dim result As Integer
    result = MsgBox("Perbaharui data ini?", vbOKCancel, "Konfirmasi")
    If result = 2 Then
        Call cmdCancel_Click
    Else
    If Not Text4.Text = "" Then
        Adodc4.Recordset!id_pinjaman = Combo1.Text
       Adodc4.Recordset!cicilan_ke = Text5.Text
       Adodc4.Recordset!tgl_bayar = txtTgl.Text
       Adodc4.Recordset.Update
        cmdSave.Enabled = False
        cmdDelete.Enabled = True
        cmdExit.Enabled = True
        gridCicilan.AllowUpdate = False
        
    Else
        result = MsgBox("Data Harus Lengkap", vbInformation, "Informasi")
        cmdAdd.Enabled = False
    End If
    End If
'        Else
'result = MsgBox("Cicilan sudah mencukupi!", vbInformation, "Informasi")
        '       cmdAdd.Enabled = False
    'End If
    
pesan:
    Select Case Err.Number
        Case -2147467259
        MsgBox "Data Sudah Ada!", vbCritical, "Perhatian"
        Call cmdCancel_Click
    End Select

End Sub


Private Sub Timer1_Timer()
sHari = aHari(Abs(Weekday(Date) - 1))
txtTgl.Text = Format(Date, "m/d/yyyy")
End Sub

Private Sub Form_Load()
Adodc4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc4.RecordSource = "select*from cicilan "
Adodc4.Refresh

'load tanggal
aHari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True



'Combobox u/ jenis kendaraan
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc1.RecordSource = "select*from pinjaman "
Adodc1.Refresh
Combo1.Clear
Do While Not Adodc1.Recordset.EOF = True
        Combo1.AddItem Adodc1.Recordset!id_pinjaman
        Combo1.ItemData(Combo1.NewIndex) = Adodc1.Recordset!id_pinjaman
        Adodc1.Recordset.MoveNext
 Loop
End Sub



