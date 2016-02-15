VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   Caption         =   "PT Permata Finance"
   ClientHeight    =   3210
   ClientLeft      =   7515
   ClientTop       =   4260
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4950
   Begin VB.Frame Frame1 
      Caption         =   "Login Form"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   480
         Top             =   2280
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
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton CmdLogin 
         Caption         =   "Login"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox TxtPassword 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox TxtUsername 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
TxtUsername.Text = ""
TxtPassword.Text = ""
End Sub

Private Sub CmdLogin_Click()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                              App.Path & "\database.mdb"
Adodc1.RecordSource = "select count(user) AS cnt from account where user ='" & TxtUsername.Text & "' AND passwd = '" & TxtPassword.Text & "'"
Adodc1.Refresh
If Adodc1.Recordset.Fields("cnt") = 0 Then
MsgBox " GAGAL LOGIN, USERNAME ATAU PASSWORD SALAH ", vbCritical
Else
Unload Me
HalamanUtama.Show
HalamanUtama.Enabled = True
End If
End Sub

