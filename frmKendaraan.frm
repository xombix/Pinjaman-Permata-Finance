VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmKendaraan 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdExit 
         Caption         =   "Keluar"
         Height          =   495
         Left            =   3600
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Simpan"
         Height          =   495
         Left            =   3600
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Hapus"
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Batal"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Tambah"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   360
         Top             =   2760
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
         RecordSource    =   "jenis_kendaraan"
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
      Begin MSDataGridLib.DataGrid gridKendaraan 
         Bindings        =   "frmKendaraan.frx":0000
         Height          =   2655
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   5175
         _ExtentX        =   9128
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
         Caption         =   "Harga"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Jenis Kendaraan"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmKendaraan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub kosong()
Text2.Text = ""
Text4.Text = ""
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
    Adodc1.Recordset.AddNew
    gridKendaraan.AllowUpdate = True

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
    gridKendaraan.AllowUpdate = False
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
    gridKendaraan.AllowUpdate = True
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
        Adodc1.Recordset!jenis = Text2.Text
        Adodc1.Recordset!harga = Text4.Text
       
        Adodc1.Recordset.Update
        cmdSave.Enabled = False
        cmdDelete.Enabled = True
        cmdExit.Enabled = True
        gridKendaraan.AllowUpdate = False
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




Private Sub gridKendaraan_Click()
Text2.Text = gridKendaraan.Columns(0).Text
Text4.Text = gridKendaraan.Columns(1).Text
End Sub


