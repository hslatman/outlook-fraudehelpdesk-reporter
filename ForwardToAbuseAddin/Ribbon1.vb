Imports Microsoft.Office.Tools.Ribbon
'Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports System
Imports System.Reflection


Public Class HOME

    Dim objItem As Outlook.MailItem
    Dim objMsg As Outlook.MailItem
    Dim app As Outlook.Application
    Dim exp As Outlook.Explorer
    Dim sel As Outlook.Selection
    Dim Application As Outlook.Application
    Dim attachments As Outlook.Attachments
    Dim objOutlookAtt As Outlook.Attachment


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles PHISHING.Click

        Dim a = Assembly.GetExecutingAssembly()

                Dim exp As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()
                If exp.Selection.Count Then
            Dim response = MsgBox("Het geselecteerde bericht zal doorgestuurd worden naar valse-email@fraudehelpdesk.nl." & vbCrLf & vbCrLf & "Wilt u doorgaan?", MsgBoxStyle.YesNo, "Fraudehelpdesk Reporter")
                    If response = MsgBoxResult.Yes Then
                        Dim selectedMail As Outlook.MailItem = exp.Selection(1)
                        Dim newMail As Outlook.MailItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem)
                        newMail.Attachments.Add(selectedMail, Outlook.OlAttachmentType.olEmbeddeditem)
                newMail.Subject = "[SPAM/PHISHING/MALWARE] - Fraudehelpdesk Reporter v" & a.GetName().Version.ToString()
                newMail.To = "hermanslatman@hotmail.com"
                        newMail.Send()
                'selectedMail.Delete()
                    Else
                    End If
                Else
            MsgBox("Selecteer alsublieft een bericht om te rapporteren.", MsgBoxStyle.OkOnly, "Fraudehelpdesk Reporter - Geen bericht geselecteerd")
                End If

    End Sub

End Class




