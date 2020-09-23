VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   7710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Scratch size"
      Height          =   2790
      Left            =   5685
      TabIndex        =   1
      Top             =   210
      Width           =   1905
      Begin VB.OptionButton Option1 
         Caption         =   "Giant"
         Height          =   270
         Index           =   6
         Left            =   90
         TabIndex        =   8
         Top             =   2265
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Biggest"
         Height          =   270
         Index           =   5
         Left            =   90
         TabIndex        =   7
         Top             =   1980
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bigger"
         Height          =   270
         Index           =   4
         Left            =   90
         TabIndex        =   6
         Top             =   1665
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Big"
         Height          =   270
         Index           =   3
         Left            =   90
         TabIndex        =   5
         Top             =   1350
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         Height          =   270
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   1035
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Small"
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   675
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Very small"
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   360
         Width           =   1140
      End
   End
   Begin ScratchAndWin.LuckyTicket LuckyTicket 
      Height          =   3000
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   5292
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Simply enjoy and invent yourself funny applications !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   9
      Top             =   3300
      Width           =   4485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
LuckyTicket.ScratchWidth = 8
Set LuckyTicket.ImageToReveal = LoadPicture(App.Path & "\Pictures\" & "ImageToBeRevealed.bmp", 4)
Set LuckyTicket.ImageToScratch = LoadPicture(App.Path & "\Pictures\" & "ImageToScratch.bmp", 4)
End Sub

Private Sub LuckyTicket1_GotFocus()

End Sub

Private Sub Option1_Click(Index As Integer)
LuckyTicket.ScratchWidth = 2 ^ (Index - 1)
End Sub
