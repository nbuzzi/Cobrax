VERSION 5.00
Begin VB.Form INICIO 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   9135
   ClientLeft      =   8925
   ClientTop       =   3450
   ClientWidth     =   13470
   LinkTopic       =   "Form4"
   Picture         =   "INICIO.frx":0000
   ScaleHeight     =   9135
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "INICIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean

Dim aniCursor1 As clsAniCursor

Private Sub Timer1_Timer()
Form1.show
Unload Me
End Sub
