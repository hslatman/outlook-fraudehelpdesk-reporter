'PhishReporter Outlook Add-In is an Outlook Add-In to Report Phishing emails to specific email addresses
'Copyright (C) 2015  Josh Rickard
'
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>
Imports Microsoft.Office.Tools.Ribbon
Imports System.Reflection

Public Class ReadMessageRibbon

    Dim objItem As Outlook.MailItem
    Dim objMsg As Outlook.MailItem
    Dim app As Outlook.Application
    Dim exp As Outlook.Explorer
    Dim sel As Outlook.Selection
    Dim Application As Outlook.Application
    Dim attachments As Outlook.Attachments
    Dim objOutlookAtt As Outlook.Attachment


    Private Sub ReadMessageRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub



    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Phishing.Click

        Dim a = Assembly.GetExecutingAssembly()
        Dim exp As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()

        If exp.Selection.Count Then
            Dim response = MsgBox("Het geselecteerde bericht zal doorgestuurd worden naar valse-email@fraudehelpdesk.nl en verplaatst worden naar verwijderde items." & vbCrLf & vbCrLf & "Wilt u doorgaan?", MsgBoxStyle.YesNo, "Fraudehelpdesk Reporter")
            If response = MsgBoxResult.Yes Then
                ' TODO: Be able to handle multiple selected messages rather than just the first one.
                Dim phishEmail As Outlook.MailItem = exp.Selection(1)
                Dim reportEmail As Outlook.MailItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem)

                reportEmail.Attachments.Add(phishEmail, Outlook.OlAttachmentType.olEmbeddeditem)
                reportEmail.Subject = "[SPAM/PHISHING/MALWARE] - Fraudehelpdesk Reporter v" & a.GetName().Version.ToString()
                reportEmail.To = "valse-email@fraudehelpdesk.nl"
                'reportEmail.Body = "This is a user-submitted report of a phishing email delivered by the PhishReporter Outlook plugin. Please review the attached phishing email"

                'If String.IsNullOrEmpty(PhishReporterConfig.RunbookURL) Then
                'reportEmail.Body = reportEmail.Body & "."
                'Else
                'reportEmail.Body = reportEmail.Body & "according to the process defined in " & PhishReporterConfig.RunbookURL
                'End If

                reportEmail.Send()
                phishEmail.Delete()
            Else
            End If
        Else
            MsgBox("Selecteer alsublieft een bericht om te rapporteren.", MsgBoxStyle.OkOnly, "Fraudehelpdesk Reporter - Geen bericht geselecteerd")
        End If

    End Sub
End Class