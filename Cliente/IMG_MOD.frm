VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Foto 
   Caption         =   $"IMG_MOD.frx":0000
   ClientHeight    =   12165
   ClientLeft      =   8220
   ClientTop       =   2520
   ClientWidth     =   15615
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   12165
   ScaleWidth      =   15615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Monitor 
      Interval        =   1
      Left            =   13080
      Top             =   3600
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&VER"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Captura simple"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6960
      OleObjectBlob   =   "IMG_MOD.frx":008D
      Top             =   9720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   255
      Left            =   14040
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   11580
      Left            =   120
      Picture         =   "IMG_MOD.frx":02C1
      ScaleHeight     =   11520
      ScaleWidth      =   15360
      TabIndex        =   0
      Top             =   480
      Width           =   15420
   End
   Begin VB.Label bytes 
      Height          =   255
      Left            =   14640
      TabIndex        =   6
      Top             =   9480
      Width           =   15
   End
   Begin VB.Label size 
      Height          =   255
      Left            =   14040
      TabIndex        =   5
      Top             =   10080
      Visible         =   0   'False
      Width           =   15
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
Dim tiempo As Long

Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer


Private Sub Command1_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    Form1.Winsock1(vIndex(0)).SendData "ssr"
    
Else
    lista.List1.ListItems.Add , , "[ Error se ha desconectado o perdido la conexion ]"
    
    Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    Unload Form2
    
ElseIf Form1.Winsock1(vIndex(0)).state = sckConnected And x = vbNo Then
    Foto.SetFocus
Else

    lista.List1.ListItems.Add , , "[ Error se ha desconectado o perdido la conexion ]"
    
    Unload Form2
    Unload Form4
    Unload Me
End If

End Sub

Private Sub Command4_Click()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    If Val(size) = Val(bytes) Then
        Form1.Winsock1(vIndex(0)).SendData "web"
    End If
Else

    Unload Form2
    Unload Form4
    Unload Me
End If

End Sub

Private Sub Form_Activate()

Foto.SetFocus
Picture1.Picture = Nothing

'Form1.Skin1.LoadSkin App.path & "\sk\skin5.skn"
'Form1.Skin1.ApplySkin Foto.hwnd

Unload fmanager
Unload procesos
End Sub
Private Sub Form_Resize()
    Dim ScaleFactorX As Single, ScaleFactorY As Single

      If Not DoResize Then
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height
      MyForm.Width = Me.Width
End Sub

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")

items_totales = (Form1.LV.ListItems.Count)

If PathFileExists(App.path & "\sk\skin.skn") Then
    Form1.Skin1.LoadSkin App.path & "\sk\skin.skn"
    Form1.Skin1.ApplySkin Form1.hwnd
End If

Dim ScaleFactorX As Single, ScaleFactorY As Single

DesignX = 800
DesignY = 600
RePosForm = True
DoResize = False

Xtwips = Screen.TwipsPerPixelX
Ytwips = Screen.TwipsPerPixelY
Ypixels = Screen.Height / Ytwips
Xpixels = Screen.Width / Xtwips

ScaleFactorX = ((Xpixels / DesignX) / 2) - 100
ScaleFactorY = ((Ypixels / DesignY) / 2) - 100
ScaleMode = 1

Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
MyForm.Height = Me.Height
MyForm.Width = Me.Width

End Sub

Private Sub Timer1_Timer()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    If Val(size) = Val(bytes) Then
        Form1.Winsock1(vIndex(0)).SendData "ssr"
    End If
Else

    lista.List1.ListItems.Add , , "[ Error se ha desconectado o perdido la conexion ]"
    Unload Form2
    Unload Form4
    Unload Me
End If


End Sub

Private Sub Monitor_Timer()
    If Check1.Value = 1 Then
    
        If Form1.Winsock1(vIndex(0)).state = sckConnected Then
        
            Form1.Winsock1(vIndex(0)).SendData "ssr"
        Else
        
            lista.List1.ListItems.Add , , "[ Error se ha perdido la conexion ]"
            
            Unload Form2
            Unload Form4
            Unload Me
        End If
    End If
End Sub

Private Sub Picture1_Click()

End Sub
