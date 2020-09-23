VERSION 5.00
Begin VB.UserControl LuckyTicket 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   Begin VB.PictureBox pctCover 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   1230
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   362
      TabIndex        =   0
      Top             =   810
      Width           =   5430
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   3045
      Left            =   720
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   362
      TabIndex        =   1
      Top             =   3000
      Width           =   5430
   End
End
Attribute VB_Name = "LuckyTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dedicated to Nicolò, Simone e Tommaso

Dim m_ScratchWidth As Integer
Const m_def_ScratchWidth = 1

Private Sub pctCover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If Button = 1 Then 'Yeah, scratching !
   pctCover.PaintPicture UserControl.Picture, x - m_ScratchWidth / 2, y - m_ScratchWidth / 2, m_ScratchWidth, m_ScratchWidth, x - m_ScratchWidth / 2, y - m_ScratchWidth / 2, m_ScratchWidth, m_ScratchWidth, vbSrcCopy
End If

End Sub

Private Sub UserControl_Initialize()
pctCover.Left = 0
pctCover.Top = 0

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   pctCover.MousePointer = 99
   pctCover.MouseIcon = LoadPicture(App.Path & "\" & "Euro.ico")
   m_ScratchWidth = PropBag.ReadProperty("ScratchWidth", m_def_ScratchWidth)
   Set Picture = PropBag.ReadProperty("ImageToReveal", Nothing)
   Set Picture = PropBag.ReadProperty("ImageToScratch", Nothing)
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 5370
UserControl.Height = 3000
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("ImageToReveal", Picture, Nothing)
   Call PropBag.WriteProperty("ImageToScratch", Picture, Nothing)
   Call PropBag.WriteProperty("ScratchWidth", m_ScratchWidth, m_def_ScratchWidth)
End Sub

'MappingInfo=pctCover,pctCover,-1,Picture
Public Property Get ImageToScratch() As Picture
Attribute ImageToScratch.VB_Description = "Restituisce o imposta un elemento grafico da visualizzare in un controllo."
   Set ImageToScratch = pctCover.Picture
End Property

Public Property Set ImageToScratch(ByVal New_ImageToScratch As Picture)
   Set pctCover.Picture = New_ImageToScratch
   PropertyChanged "ImageToScratch"
End Property

'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get ImageToReveal() As Picture
Attribute ImageToReveal.VB_Description = "Restituisce o imposta un elemento grafico da visualizzare in un controllo."
   Set ImageToReveal = UserControl.Picture
End Property

Public Property Set ImageToReveal(ByVal New_ImageToReveal As Picture)
   Set UserControl.Picture = New_ImageToReveal
   PropertyChanged "ImageToReveal"
End Property

'MemberInfo=7,0,0,1
Public Property Get ScratchWidth() As Integer
   ScratchWidth = m_ScratchWidth
End Property

Public Property Let ScratchWidth(ByVal New_ScratchWidth As Integer)
   m_ScratchWidth = New_ScratchWidth
   PropertyChanged "ScratchWidth"
End Property

Private Sub UserControl_InitProperties()
   m_ScratchWidth = m_def_ScratchWidth
End Sub

