Partial Class ReadMessageRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReadMessageRibbon))
        Me.PhishReporter_Read_Message = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Phishing = Me.Factory.CreateRibbonButton
        Me.PhishReporter_Read_Message.SuspendLayout()
        Me.Group2.SuspendLayout()
        '
        'PhishReporter_Read_Message
        '
        Me.PhishReporter_Read_Message.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.PhishReporter_Read_Message.ControlId.OfficeId = "TabReadMessage"
        Me.PhishReporter_Read_Message.Groups.Add(Me.Group2)
        Me.PhishReporter_Read_Message.Label = "TabReadMessage"
        Me.PhishReporter_Read_Message.Name = "PhishReporter_Read_Message"
        Me.PhishReporter_Read_Message.Position = Me.Factory.RibbonPosition.BeforeOfficeId("GroupQuickSteps")
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Phishing)
        Me.Group2.Label = "Fraudehelpdesk"
        Me.Group2.Name = "Group2"
        Me.Group2.Position = Me.Factory.RibbonPosition.BeforeOfficeId("GroupQuickSteps")
        '
        'Phishing
        '
        Me.Phishing.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Phishing.Image = CType(resources.GetObject("Phishing.Image"), System.Drawing.Image)
        Me.Phishing.Label = "Rapporteren"
        Me.Phishing.Name = "Phishing"
        Me.Phishing.OfficeImageId = "TrustCenter"
        Me.Phishing.ScreenTip = "Rapporteer email"
        Me.Phishing.ShowImage = True
        Me.Phishing.SuperTip = "Stuur deze email door naar de Fraudehelpdesk"
        '
        'ReadMessageRibbon
        '
        Me.Name = "ReadMessageRibbon"
        Me.RibbonType = "Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.PhishReporter_Read_Message)
        Me.PhishReporter_Read_Message.ResumeLayout(False)
        Me.PhishReporter_Read_Message.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()

    End Sub

    Friend WithEvents PhishReporter_Read_Message As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Phishing As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property ReadMessageRibbon() As ReadMessageRibbon
        Get
            Return Me.GetRibbon(Of ReadMessageRibbon)()
        End Get
    End Property
End Class