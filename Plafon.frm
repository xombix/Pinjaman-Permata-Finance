VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPlafon 
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   5655
   ClientTop       =   3795
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   6645
   Begin VB.Frame Frame1 
      Caption         =   "Simulasi Plafon"
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "Hitung"
         Height          =   495
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Plafon.frx":0000
         Left            =   1560
         List            =   "Plafon.frx":0002
         TabIndex        =   6
         Top             =   1320
         Width           =   2295
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   120
         Top             =   3240
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   1560
         Top             =   3240
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
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Cicilan Perbulan : "
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "Pinjaman"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   6255
      End
      Begin VB.Label Label3 
         Caption         =   "Tenor"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Tahun Kendaraan"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Jenis Kendaraan"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmPlafon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
'Label4.Caption = Combo3.List(Combo3.ListIndex)
Label4.Caption = "Total Pinjaman :  Rp. " & Format(Val(Combo2.ItemData(Combo2.ListIndex)) + Val(Combo1.ItemData(Combo1.ListIndex)), "###,###,##0.00")
 total = Combo2.ItemData(Combo2.ListIndex) + Combo1.ItemData(Combo1.ListIndex)
 Gtotal = (Val(total) / Val(Combo3.ItemData(Combo3.ListIndex))) + (0.05 * Val(total))
 Label5.Caption = "Jumlah Angsuran Perbulan :  Rp. " & Format(Round(Gtotal), "###,###,##0.00")
End Sub

Private Sub Form_Load()
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

 
End Sub

