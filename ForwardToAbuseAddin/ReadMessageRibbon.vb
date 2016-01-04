Imports Microsoft.Office.Tools.Ribbon
Imports System.Reflection

Public Class ReadMessageRibbon

    Private Sub ReadMessageRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub


    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Phishing.Click

        ClickCode.Click()

    End Sub
End Class