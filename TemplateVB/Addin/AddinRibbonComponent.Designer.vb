Partial Class AddinRibbonComponent
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
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Command1 = Me.Factory.CreateRibbonButton
        $if$ ($taskpane$ == true)Me.TogglePanel = Me.Factory.CreateRibbonButton
        $endif$Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.ControlId.OfficeId = "TabHome"
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabHome"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Command1)
        $if$ ($taskpane$ == true)Me.Group1.Items.Add(Me.TogglePanel)
        $endif$Me.Group1.Label = $productNameVB$
        Me.Group1.Name = "Group1"
        '
        'Command1
        '
        Me.Command1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Command1.Image = Global.$csprojectname$.My.Resources.Resources.Command1
        Me.Command1.Label = "Command 1"
        Me.Command1.Name = "Command1"
        Me.Command1.ShowImage = True
        $if$ ($taskpane$ == true)
        '
        'TogglePanel
        '
        Me.TogglePanel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.TogglePanel.Image = Global.$csprojectname$.My.Resources.Resources.TogglePanel
        Me.TogglePanel.Label = "Toggle Panel"
        Me.TogglePanel.Name = "TogglePanel"
        Me.TogglePanel.ShowImage = True
        $endif$
        '
        'AddinRibbonComponent
        '
        Me.Name = "AddinRibbonComponent"
        Me.RibbonType = "Microsoft.Visio.Drawing"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Command1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    $if$ ($taskpane$ == true)Friend WithEvents TogglePanel As Microsoft.Office.Tools.Ribbon.RibbonButton
$endif$End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As AddinRibbonComponent
        Get
            Return Me.GetRibbon(Of AddinRibbonComponent)()
        End Get
    End Property
End Class
