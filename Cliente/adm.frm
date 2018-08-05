VERSION 5.00
Begin VB.Form adm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrador de tareas (Remoto)"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5865
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Detener"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar proceso"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton Actualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   8400
      Width           =   1215
   End
   Begin VB.ListBox procesos 
      Height          =   7860
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "adm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Actualizar_Click()
Form1.Winsock1(vIndex).SendData "lis"
listaproceso = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
listaproceso = False
End Sub

Private Sub Form_Load()
listaproceso = True
End Sub
