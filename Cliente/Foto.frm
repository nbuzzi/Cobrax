VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Foto 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Monitoreo de sistema (Ordenador) - BETA 1"
   ClientHeight    =   12660
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   15810
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12660
   ScaleWidth      =   15810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6960
      OleObjectBlob   =   "Foto.frx":0000
      Top             =   12000
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   11640
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   12120
      TabIndex        =   2
      Top             =   11640
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recibir Imagen"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   11640
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   11295
      Left            =   120
      Picture         =   "Foto.frx":0234
      ScaleHeight     =   11235
      ScaleWidth      =   15435
      TabIndex        =   0
      Top             =   120
      Width           =   15495
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   0
         Top             =   6600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   101
         LocalPort       =   101
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Datos recibidos actualmente : %0 kb's"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   12240
      Width           =   6495
   End
End
Attribute VB_Name = "Foto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nombre As String

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpszExistingFileName As String) As Boolean
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lpszExistingPath As String) As Boolean

Private Sub Command1_Click()
Form1.Winsock1(vIndex).SendData "ssr"
Winsock1.Listen
End Sub

Private Sub Command2_Click()
Winsock2.Close
Unload Me
End Sub

Private Sub Form_Activate()
Skin1.LoadSkin App.path & "\sk\skin5.skn"
Skin1.ApplySkin Foto.hwnd
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
    Winsock2.Close
    Winsock2.Accept requestID
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)

'Array dinámico para almacenar el archivo
Dim imagen() As Byte


    nombre = Str(Rnd(5000) & Rnd(5000) & ".bmp")
    Open App.path + "\" & nombre For Binary As #1

    'Esto escribe en el disco el archivo de imagen
    Winsock2.GetData imagen
    If UBound(imagen) = 1 Then
        Close
        Picture1 = LoadPicture(App.path + "\" + nombre)

        Kill App.path + "\" + nombre
    End If
               
    'Escribimos en disco la imagen con Put pasandole el Array
    Put #1, , imagen
    
End Sub
