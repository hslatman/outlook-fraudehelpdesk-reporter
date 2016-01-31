Imports System.Reflection
Imports System.Xml
Imports System.Net
Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms
Imports System.Configuration

Module ForwardCode
    Sub Click()

        Dim a = Assembly.GetExecutingAssembly()
        Dim exp As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()

        If exp.Selection.Count Then
            'Dim response = MsgBox("Het geselecteerde bericht zal doorgestuurd worden naar valse-email@fraudehelpdesk.nl en verplaatst worden naar verwijderde items." & vbCrLf & vbCrLf & "Wilt u doorgaan?", MsgBoxStyle.YesNo, "Fraudehelpdesk Reporter")
            Dim frm As Form1 = New Form1()
            frm.ShowDialog()

            If frm.DialogResult = DialogResult.Yes Then

                frm.Dispose()

                'TODO: handle multiple selected messages rather than just the first one.
                Dim phishEmail As Outlook.MailItem = exp.Selection(1)
                Dim reportEmail As Outlook.MailItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem)

                reportEmail.Attachments.Add(phishEmail, Outlook.OlAttachmentType.olEmbeddeditem)
                reportEmail.Subject = "[SPAM/PHISHING/MALWARE] - Fraudehelpdesk Reporter v" & a.GetName().Version.ToString()
                reportEmail.To = MySettings.Default.ReportAddress
                reportEmail.Body = "Deze email is verstuurd met de Fraudehelpdesk Reporter."

                Dim pa As Outlook.PropertyAccessor = reportEmail.PropertyAccessor
                pa.SetProperty("http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-FHD-Reporter-Version", a.GetName().Version.ToString())
                pa.SetProperty("http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-FHD-Reporter-Reaction", MySettings.Default.SendReaction.ToString())
                pa.SetProperty("http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-FHD-Reporter-Updates", MySettings.Default.CheckUpdates.ToString())

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

    Sub Update()

        'Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        'Dim userSettings As UserSettingsGroup = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).GetSection()


        'Dim configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        'Dim checkUpdate As ClientSettingsSection = configFile.GetSection("userSettings/ForwardToAbuseAddin.MySettings")

        'Check the settings file if we should check for updates.
        'Dim shouldCheck As Boolean = Convert.ToBoolean(checkUpdate.Settings.Get("CheckUpdates").Value.ValueXml.InnerText.ToString())

        If (MySettings.Default.CheckUpdates) Then

            'check the current assembly version
            Dim a = Assembly.GetExecutingAssembly()
            Dim version = a.GetName().Version

            Dim URLString As String = "http://localhost/fhd_download_manifest.xml"
            Dim wrGETURL As WebRequest
            wrGETURL = WebRequest.Create(URLString)
            wrGETURL.Proxy = WebRequest.DefaultWebProxy()
            Dim objStream As Stream

            Try
                'Try to connect to the manifest download URL
                objStream = wrGETURL.GetResponse.GetResponseStream()


                Dim doc As XmlDocument = New XmlDocument()
                doc.Load(objStream)

                Dim v As String = ""

                Try
                    'Try to get the version as a String
                    v = doc.SelectSingleNode("fhd/@version").InnerText
                Catch ex As Exception
                    'Set a lower version
                    v = "1.0.0.0"
                End Try

                Try
                    'Only continue if we have an URL attribute
                    Dim url As String = doc.SelectSingleNode("fhd/@url").InnerText

                    'Create a 'real' Assembly.Version out of it, for comparison
                    Dim ManifestVersion = New Version(v)

                    Dim result = version.CompareTo(ManifestVersion)
                    'result > 0, then we don't have a new version...
                    'result < 0, then the ManifestVersion is bigger, so we have to download the new one...

                    If (result < 0) Then
                        Dim response = MsgBox("Een nieuwe versie van de Fraudehelpdesk Reporter Outlook Add-In is beschikbaar." & vbCrLf & vbCrLf & "Wilt u deze nu downloaden en installeren?", MsgBoxStyle.YesNo, "Fraudehelpdesk Reporter")
                        If response = MsgBoxResult.Yes Then

                            'Start the download by navigating via the default browser
                            'Code taken from: http://code.logos.com/blog/2008/01/using_processstart_to_link_to.html
                            Try
                                Process.Start(url)

                            Catch exception As Exception


                            Finally

                            End Try

                        End If

                    End If

                Catch ex As Exception

                End Try

            Catch exception As Exception

                'No connection to the URL could be made

            Finally

            End Try


        End If
        'We don't have to check for updates...

    End Sub
End Module
