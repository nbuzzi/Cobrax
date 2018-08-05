VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form bromas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bromas al usuario"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4575
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Pintar ventana actual"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Dibujar circulos en el escritorio"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Dibujar escritorio"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1680
      OleObjectBlob   =   "bromas.frx":0000
      Top             =   3000
   End
End
Attribute VB_Name = "bromas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vIndex As Variant
Dim Index_n As Integer
Dim x As Integer

Private Sub Command1_Click()
x = MsgBox("Esta funcion durara 1 minutos, durante ese tiempo no podra ingresar comandos." & vbCrLf & "Esta seguro que desea hacerla?", vbInformation + vbYesNo, "Aviso")
If x = vbYes Then
    If Form1.Winsock1(vIndex(0)).state = sckConnected Then
        Form1.Winsock1(vIndex(0)).SendData "xam"
    Else
        lista.List1.ListItems.Add , , "[Error al procesar comando - conexión perdida]"
        Unload Me
    End If
End If
End Sub

Private Sub Command2_Click()
x = MsgBox("Esta funcion durara 1 minutos, durante ese tiempo no podra ingresar comandos." & vbCrLf & "Esta seguro que desea hacerla?", vbInformation + vbYesNo, "Aviso")
If x = vbYes Then
    If Form1.Winsock1(vIndex(0)).state = sckConnected Then
        Form1.Winsock1(vIndex(0)).SendData "xac"
    Else
        lista.List1.ListItems.Add , , "[Error al procesar comando - conexión perdida]"
        Unload Me
    End If
End If
End Sub

Private Sub Command3_Click()
x = MsgBox("Esta funcion durara 1 minutos, durante ese tiempo no podra ingresar comandos." & vbCrLf & "Esta seguro que desea hacerla?", vbInformation + vbYesNo, "Aviso")
If x = vbYes Then
    If Form1.Winsock1(vIndex(0)).state = sckConnected Then
        Form1.Winsock1(vIndex(0)).SendData "xaz"
    Else
        lista.List1.ListItems.Add , , "[Error al procesar comando - conexión perdida]"
        Unload Me
    End If
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")


Skin1.LoadSkin App.path & "\sk\skin.skn"
Skin1.ApplySkin ventanas.hwnd
End Sub
