Attribute VB_Name = "Module1"
Option Explicit
  
Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long
  
Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long
  
Dim texto As String
Dim mForm As Form
Public mEspacios As Integer
  
  
Sub Comenzar(frm As Form, Intervalo As Long, mtexto As String)
    Call Detener(frm)
    frm.Caption = ""
    mEspacios = (frm.ScaleWidth) / 70
    texto = mtexto
    Set mForm = frm
    SetTimer frm.hwnd, 0, Intervalo, AddressOf TimerProc
End Sub
  
Sub Detener(frm As Form)
  
    KillTimer frm.hwnd, 0
    Set mForm = Nothing
End Sub

Sub TimerProc(ByVal hwnd As Long, _
              ByVal nIDEvent As Long, _
              ByVal uElapse As Long, _
              ByVal lpTimerFunc As Long)
      
    Call scroll
  
End Sub
  
  
Public Sub scroll()
    If Not mForm.Caption = "" Then
        mForm.Caption = Right(mForm.Caption, (Len(mForm.Caption) - 1))
    Else
        mForm.Caption = Space(mEspacios)
        mForm.Caption = mForm.Caption + texto
    End If
End Sub

