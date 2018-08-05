VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F8180939-60A2-4494-B1BB-04818D7F640B}#1.0#0"; "LabelDegradado.ocx"
Begin VB.Form Configuracion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuación del cliente"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3285
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LabelDegradado.LabelDegrade LabelDegrade2 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      Text            =   "Puerto de escucha"
      BackColor       =   255
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      ShadowColorStart=   0
      DegradadoColorStart=   8388608
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1200
      OleObjectBlob   =   "Configuracion.frx":0000
      Top             =   0
   End
   Begin LabelDegradado.LabelDegrade LabelDegrade1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Aqui entraran las conexiones"
      BackColor       =   16711680
      Transparente    =   0   'False
      ShadowDepth     =   0
      ShadowStyle     =   0
      ShadowColorStart=   0
      DegradadoColorStart=   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "100"
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If IsNumeric(Text1) Then

    Open App.path & "\Config.ini" For Output As #1
    Print #1, "[USERCONFIG]" & vbCrLf & Val(Text1) & vbCrLf & "defaultuser=" & Environ("computername") & vbCrLf
    Close #1
    
    x = MsgBox("Debe reiniciar el programa para que los cambios sean aplicados, ¿ Desea cerrar el programa ahora ?", vbInformation + vbYesNo, "Aviso")
    
    If x = vbYes Then
       End
    End If
End If

Unload Me

End Sub

Private Sub Form_Load()
Skin1.LoadSkin App.path & "\sk\skin5.skn"
Skin1.ApplySkin Configuracion.hwnd
End Sub
