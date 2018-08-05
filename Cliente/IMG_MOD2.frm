VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Foto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitoreo de sistema (Ordenador) - BETA 1"
   ClientHeight    =   10665
   ClientLeft      =   -345
   ClientTop       =   4995
   ClientWidth     =   14730
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10665
   ScaleWidth      =   14730
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   13920
      TabIndex        =   6
      Text            =   "0"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Ampliar"
      Height          =   375
      Left            =   13920
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13920
      Top             =   2520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "VER"
      Height          =   255
      Left            =   13920
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Recibir"
      Height          =   375
      Left            =   13920
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6960
      OleObjectBlob   =   "IMG_MOD2.frx":0000
      Top             =   9720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   13920
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   10500
      Left            =   120
      Picture         =   "IMG_MOD2.frx":0234
      ScaleHeight     =   10440
      ScaleWidth      =   13635
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   16920
         Top             =   3960
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Datos recibidos actualmente : %0 kb's"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   11760
      Width           =   6495
   End
End
Attribute VB_Name = "Foto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpszExistingFileName As String) As Boolean
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lpszExistingPath As String) As Boolean
Dim vIndex As Variant
Dim Index_n As Integer
Dim items_totales As Integer

Private Sub Command1_Click()

If Form1.Winsock1(vIndex(0)).State = sckConnected Then

    Form1.Winsock1(vIndex(0)).SendData "ssr"
    
Else
    Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Check1.Value = 0 And PathFileExists(App.path & "\imagen.jpg") Then
Shell App.path & "\imagen.jpg", vbMaximizedFocus
End If
End Sub

Private Sub Form_Activate()

Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")

items_totales = (Form1.LV.ListItems.Count)

Foto.SetFocus
Picture1.Picture = Nothing

Form1.Skin1.LoadSkin App.path & "\sk\skin5.skn"
Form1.Skin1.ApplySkin Foto.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
If PathFileExists(App.path & "\sk\skin.skn") Then
    Form1.Skin1.LoadSkin App.path & "\sk\skin.skn"
    Form1.Skin1.ApplySkin Form1.hwnd
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)

If Form1.Winsock1(vIndex(0)).State = sckConnected Then
    If KeyAscii = 13 And IsNumeric(Text2.Text) Then
        
    Check1.Value = 1
    Text2.Enabled = False
        
    ElseIf Not IsNumeric(Text2.Text) Then
        Check1.Value = 0
        Text2.Text = 0
        Text2.SetFocus
    End If
Else
    Unload Form2
    Unload Form4
    Unload Me
    lista.List1.AddItem "[Error al intentar usar el monitoreo - conexion perdida]"
End If

End Sub

Private Sub Timer1_Timer()

If Form1.Winsock1(vIndex(0)).State = sckConnected Then

    Form1.Winsock1(vIndex(0)).SendData "ssr"
    
Else
    Unload Me
End If



End Sub

Private Sub Timer2_Timer()

If Check1.Value = 1 Then
    Timer1.Enabled = True
    Text2.Enabled = False
Else
    Text2.Enabled = True
    Timer1.Enabled = False
End If

If Form1.LV.SelectedItem.Index > 0 Then
    Index_n = (Form1.LV.SelectedItem.Index)
    vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")
End If

If Form1.Winsock1(vIndex(0)).State = sckError Then
    Unload Form2
    Unload Form4
    Unload Me
End If

If Not IsNumeric(Text2.Text) Then
Text2.Text = 1
End If

If Text2.Text = 0 Then
    Timer1.Interval = 100
Else
    Timer1.Interval = Text2.Text * 1000
End If
End Sub

