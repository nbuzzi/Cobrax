Attribute VB_Name = "Module3"
Public Xtwips As Integer, Ytwips As Integer
Public Xpixels As Integer, Ypixels As Integer

    Type FRMSIZE
    Height As Long
    Width As Long
End Type

Public RePosForm As Boolean
Public DoResize As Boolean

Sub Resize_For_Resolution(ByVal SFX As Single, _
    ByVal SFY As Single, MyForm As Form)
    Dim I As Integer
    Dim SFFont As Single

    SFFont = (SFX + SFY) / 2

    On Error Resume Next
    With MyForm
        For I = 0 To .Count - 1
            If TypeOf .Controls(I) Is ComboBox Then
                .Controls(I).Left = .Controls(I).Left * SFX
                .Controls(I).Top = .Controls(I).Top * SFY
                .Controls(I).Width = .Controls(I).Width * SFX
            Else
                .Controls(I).Move .Controls(I).Left * SFX, _
                .Controls(I).Top * SFY, _
                .Controls(I).Width * SFX, _
                .Controls(I).Height * SFY
            End If

                .Controls(I).FontSize = .Controls(I).FontSize * SFFont
            Next I
        If RePosForm Then
          .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
        End If
    End With
End Sub

