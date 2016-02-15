VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form lapNasabah 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Submit"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker tgl2 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   100663297
         CurrentDate     =   42405
      End
      Begin MSComCtl2.DTPicker tgl1 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   100663297
         CurrentDate     =   42405
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal Akhir"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal Awal"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "lapNasabah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
