VERSION 5.00
Begin VB.MDIForm HalamanUtama 
   BackColor       =   &H80000004&
   Caption         =   "Permata Finance"
   ClientHeight    =   6255
   ClientLeft      =   6285
   ClientTop       =   3090
   ClientWidth     =   6630
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   6135
      Left            =   0
      Picture         =   "HalamanUtama.frx":0000
      ScaleHeight     =   6075
      ScaleWidth      =   6570
      TabIndex        =   0
      Top             =   0
      Width           =   6630
      Begin VB.Timer Timer1 
         Left            =   2400
         Top             =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3000
         TabIndex        =   1
         Top             =   1200
         Width           =   3375
      End
   End
   Begin VB.Menu MData 
      Caption         =   "Master Data"
      Begin VB.Menu surveyor 
         Caption         =   "Surveyor"
      End
      Begin VB.Menu account 
         Caption         =   "Account"
      End
      Begin VB.Menu jenis 
         Caption         =   "Jenis Kendaraan"
      End
      Begin VB.Menu tahun 
         Caption         =   "Tahun Kendaraan"
      End
      Begin VB.Menu nasabah 
         Caption         =   "Nasabah"
      End
   End
   Begin VB.Menu simulasi 
      Caption         =   "Simulasi"
      Begin VB.Menu SPlafon 
         Caption         =   "Simulasi Plafon"
      End
   End
   Begin VB.Menu transaksi1 
      Caption         =   "Transaksi"
      Begin VB.Menu PPinjaman 
         Caption         =   "Pengajuan Pinjaman"
      End
      Begin VB.Menu BCicilan 
         Caption         =   "Bayar Cicilan"
      End
   End
   Begin VB.Menu laporan 
      Caption         =   "Laporan"
      Begin VB.Menu LNasabah 
         Caption         =   "Laporan Nasabah"
      End
      Begin VB.Menu LPinjaman 
         Caption         =   "Laporan Pinjaman"
      End
   End
   Begin VB.Menu keluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "HalamanUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub account_Click()
frmAccount.Show
End Sub

Private Sub BCicilan_Click()
frmCicilan.Show
End Sub

Private Sub jenis_Click()
frmKendaraan.Show
End Sub

Private Sub keluar_Click()
End
End Sub

Private Sub LNasabah_Click()
rptNasabah.Show
End Sub

Private Sub LPinjaman_Click()
rptPinjaman.Show
End Sub

Private Sub MDIForm_Load()
aHari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True
sHari = aHari(Abs(Weekday(Date) - 1))
Label1.Caption = sHari & ", " & Format(Date, "d/m/yyyy")
End Sub

Private Sub nasabah_Click()
frmNasabah.Show
End Sub

Private Sub PPinjaman_Click()
frmPinjaman.Show
End Sub

Private Sub SPlafon_Click()
frmPlafon.Show
End Sub

Private Sub surveyor_Click()
frmSurveyor.Show
End Sub

Private Sub tahun_Click()
frmTahun.Show
End Sub

