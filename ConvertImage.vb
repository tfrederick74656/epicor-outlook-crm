Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing

Public Class ConvertImage
    Inherits System.Windows.Forms.AxHost

    Public Sub New()
        MyBase.New("59EE46BA-677D-4d20-BF10-8D8067CB8B32")
    End Sub

    Public Shared Function Convert(ByVal Image As System.Drawing.Image) As stdole.IPictureDisp
        Convert = CType(GetIPictureFromPicture(Image), stdole.IPictureDisp)
    End Function
End Class