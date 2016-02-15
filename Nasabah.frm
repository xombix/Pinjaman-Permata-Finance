VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmNasabah 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   14205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.CommandButton cmdExit 
         Caption         =   "Keluar"
         Height          =   495
         Left            =   9000
         TabIndex        =   31
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Simpan"
         Height          =   495
         Left            =   9000
         TabIndex        =   30
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   7200
         TabIndex        =   29
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   7200
         TabIndex        =   28
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Batal"
         Height          =   495
         Left            =   5640
         TabIndex        =   27
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Tambah"
         Height          =   495
         Left            =   5640
         TabIndex        =   26
         Top             =   3240
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2160
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=database.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=database.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "nasabah"
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
      Begin MSDataGridLib.DataGrid gridNasabah 
         Bindings        =   "Nasabah.frx":0000
         Height          =   2655
         Left            =   120
         TabIndex        =   25
         Top             =   4680
         Width           =   13455
         _ExtentX        =   23733
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
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   8400
         TabIndex        =   24
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   8400
         TabIndex        =   23
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   8400
         TabIndex        =   22
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   8400
         TabIndex        =   21
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   8400
         TabIndex        =   20
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   8400
         TabIndex        =   19
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   1575
         Left            =   2160
         TabIndex        =   15
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label12 
         Caption         =   "Warna Kendaraan"
         Height          =   495
         Left            =   6480
         TabIndex        =   12
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "No Mesin"
         Height          =   495
         Left            =   6480
         TabIndex        =   11
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "No Rangka "
         Height          =   495
         Left            =   6480
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "No. BPKB"
         Height          =   495
         Left            =   6480
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "No. STNK"
         Height          =   375
         Left            =   6480
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Tahun Kendaraan"
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Jenis Kendaraan"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "No. KTP"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "No Handphone"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Alamat"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Nama"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "No Pk"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmNasabah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub kosong()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text1.SetFocus
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
    Adodc1.Recordset.AddNew
    gridNasabah.AllowUpdate = True

End Sub

Private Sub cmdCancel_Click()
kosong
cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdDelete.Enabled = True
    cmdExit.Enabled = True
    Adodc1.Recordset.Cancel
    Adodc1.Refresh
    gridNasabah.AllowUpdate = False
End Sub

Private Sub cmdDelete_Click()
 Dim result As Integer
    result = MsgBox("Hapus data ini?", vbOKCancel, "Konfirmasi")
    If result = 2 Then
        Adodc1.Recordset.CancelUpdate
            Else
        If Not Adodc1.Recordset.EOF Then
            Adodc1.Recordset.Delete
            Adodc1.Recordset.MoveFirst
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
    gridNasabah.AllowUpdate = True
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
    If Not Text1.Text = "" Then
        Adodc1.Recordset!no_pk = Text1.Text
        Adodc1.Recordset!nama = Text2.Text
        Adodc1.Recordset!alamat = Text3.Text
        Adodc1.Recordset!no_hp = Text4.Text
        Adodc1.Recordset!no_ktp = Text5.Text
        Adodc1.Recordset!jenis_kendaraan = Text6.Text
        Adodc1.Recordset!tahun_kendaraan = Text7.Text
        Adodc1.Recordset!no_stnk = Text8.Text
        Adodc1.Recordset!no_bpkb = Text9.Text
        Adodc1.Recordset!no_rangka = Text10.Text
        Adodc1.Recordset!no_mesin = Text11.Text
        Adodc1.Recordset!warna_kendaraan = Text12.Text
       
        Adodc1.Recordset.Update
        cmdSave.Enabled = False
        cmdDelete.Enabled = True
        cmdExit.Enabled = True
        gridNasabah.AllowUpdate = False
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

Private Sub gridNasabah_Click()
Text1.Text = gridNasabah.Columns(0).Text
Text2.Text = gridNasabah.Columns(1).Text
Text3.Text = gridNasabah.Columns(2).Text
Text4.Text = gridNasabah.Columns(3).Text
Text5.Text = gridNasabah.Columns(4).Text
Text6.Text = gridNasabah.Columns(5).Text
Text7.Text = gridNasabah.Columns(6).Text
Text8.Text = gridNasabah.Columns(7).Text
Text9.Text = gridNasabah.Columns(8).Text
Text10.Text = gridNasabah.Columns(9).Text
Text11.Text = gridNasabah.Columns(10).Text
Text12.Text = gridNasabah.Columns(11).Text

End Sub
