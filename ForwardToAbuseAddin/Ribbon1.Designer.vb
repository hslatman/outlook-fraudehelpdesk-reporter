Partial Class HOME
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(HOME))
        Me.PhishReporter = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.PHISHING = Me.Factory.CreateRibbonButton
        Me.PhishReporter.SuspendLayout()
        Me.Group1.SuspendLayout()
        '
        'PhishReporter
        '
        Me.PhishReporter.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.PhishReporter.ControlId.OfficeId = "TabMail"
        Me.PhishReporter.Groups.Add(Me.Group1)
        Me.PhishReporter.Label = "TabMail"
        Me.PhishReporter.Name = "PhishReporter"
        Me.PhishReporter.Position = Me.Factory.RibbonPosition.BeforeOfficeId("GroupQuickSteps")
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.PHISHING)
        Me.Group1.Label = "Fraudehelpdesk"
        Me.Group1.Name = "Group1"
        Me.Group1.Position = Me.Factory.RibbonPosition.BeforeOfficeId("GroupQuickSteps")
        '
        'PHISHING
        '
        Me.PHISHING.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.PHISHING.Image = CType(resources.GetObject("PHISHING.Image"), System.Drawing.Image)
        Me.PHISHING.Label = "Rapporteren"
        Me.PHISHING.Name = "PHISHING"
        Me.PHISHING.OfficeImageId = "TrustCenter"
        Me.PHISHING.ShowImage = True
        Me.PHISHING.SuperTip = "Stuur deze email door naar de Fraudehelpdesk"
        '
        'HOME
        '
        Me.Name = "HOME"
        Me.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mai" & _
    "l.Read"
        Me.Tabs.Add(Me.PhishReporter)
        Me.PhishReporter.ResumeLayout(False)
        Me.PhishReporter.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()

    End Sub

    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents PHISHING As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents PhishReporter As Microsoft.Office.Tools.Ribbon.RibbonTab
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As HOME
        Get
            Return Me.GetRibbon(Of HOME)()
        End Get
    End Property
End Class
