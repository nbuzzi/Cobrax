VERSION 5.00
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form Form5 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   3600
   ClientLeft      =   24090
   ClientTop       =   12660
   ClientWidth     =   4800
   LinkTopic       =   "Form5"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin LabelDegradado.LabelDegrade Label1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "                 &Nueva conexión"
      BackColor       =   255
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      ShadowColorStart=   0
      DegradadoColorStart=   0
      DegradadoColorEnd=   255
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   240
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Private Sub Form_Load()
    Timer1.Interval = 1
    Timer1.Enabled = True
    Label2.Caption = "IP: " + Form1.IPRemte
    Me.Move 0, 0
End Sub

Private Sub Timer1_Timer()
    If i < 1950 Then
        i = i + 20
        Me.Width = i
    Else
        Unload Me
    End If
End Sub
