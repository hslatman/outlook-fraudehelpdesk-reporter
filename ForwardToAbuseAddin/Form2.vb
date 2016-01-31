Imports System.Configuration

Public Class Form2

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        MySettings.Default.SendReaction = CheckBox1.Checked
        'Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        'config.Save()
        'ConfigurationManager.RefreshSection("userSettings/ForwardToAbuseAddin.MySettings")
        MySettings.Default.Save()

        ConfigurationManager.RefreshSection("userSettings/ForwardToAbuseAddin.MySettings")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        MySettings.Default.CheckUpdates = CheckBox2.Checked

        MySettings.Default.Save()
        ConfigurationManager.RefreshSection("userSettings/ForwardToAbuseAddin.MySettings")

    End Sub
End Class