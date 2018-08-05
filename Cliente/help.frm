VERSION 5.00
Begin VB.Form help 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   10155
   ClientLeft      =   6690
   ClientTop       =   2235
   ClientWidth     =   13485
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "help.frx":0000
   ScaleHeight     =   10155
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "ACEPTO!"
      Height          =   495
      Left            =   11280
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   $"help.frx":4C9D8
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   8880
      Width           =   13455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   $"help.frx":4CBA5
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpCommand As String, ByVal lpreturnsa As Long, ByVal lpparameters As Long, ByVal otherparam As Long) As Long
Dim aniCursor1 As clsAniCursor

Private Sub Command1_Click()
mciSendString "stop sk\ambice.mp3", 0, 0, 0
Unload Me
End Sub

Private Sub Form_Load()
mciSendString "play sk\ambice.mp3", 0, 0, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
mciSendString "stop sk\ambice.mp3", 0, 0, 0
End Sub
