Attribute VB_Name = "Module2"
Option Explicit
  
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long
  

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                 ByVal nIndex As Long) As Long
  
  
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
  
  
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

  
Public Function Is_Transparent(ByVal hwnd As Long) As Boolean
On Error Resume Next
  
Dim msg As Long
  
    msg = GetWindowLong(hwnd, GWL_EXSTYLE)
         
       If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
          Is_Transparent = True
       Else
          Is_Transparent = False
       End If
  
    If Err Then
       Is_Transparent = False
    End If
  
End Function

Public Function Aplicar_Transparencia(ByVal hwnd As Long, _
                                      Valor As Integer) As Long
  
Dim msg As Long
  
On Error Resume Next
  
If Valor < 0 Or Valor > 255 Then
   Aplicar_Transparencia = 1
Else
   msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   msg = msg Or WS_EX_LAYERED
     
   SetWindowLong hwnd, GWL_EXSTYLE, msg
   
   SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA
  
   Aplicar_Transparencia = 0
  
End If
  
  
If Err Then
   Aplicar_Transparencia = 2
End If
  
End Function

