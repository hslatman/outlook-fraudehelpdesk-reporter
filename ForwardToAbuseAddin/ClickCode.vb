Imports System.Reflection

Module ClickCode
    Sub Click()

        Dim a = Assembly.GetExecutingAssembly()
        Dim exp As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()

        If exp.Selection.Count Then
            Dim response = MsgBox("Het geselecteerde bericht zal doorgestuurd worden naar valse-email@fraudehelpdesk.nl en verplaatst worden naar verwijderde items." & vbCrLf & vbCrLf & "Wilt u doorgaan?", MsgBoxStyle.YesNo, "Fraudehelpdesk Reporter")
            If response = MsgBoxResult.Yes Then
                'TODO: handle multiple selected messages rather than just the first one.
                Dim phishEmail As Outlook.MailItem = exp.Selection(1)
                Dim reportEmail As Outlook.MailItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem)

                reportEmail.Attachments.Add(phishEmail, Outlook.OlAttachmentType.olEmbeddeditem)
                reportEmail.Subject = "[SPAM/PHISHING/MALWARE] - Fraudehelpdesk Reporter v" & a.GetName().Version.ToString()
                reportEmail.To = "valse-email@fraudehelpdesk.nl"
                reportEmail.Body = "Deze email is verstuurd met de Fraudehelpdesk Reporter."

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
End Module
