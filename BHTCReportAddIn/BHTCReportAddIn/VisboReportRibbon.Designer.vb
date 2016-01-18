Partial Class VisboReportRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Erforderlich für die Unterstützung des Windows.Forms-Klassenkompositions-Designers
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Dieser Aufruf ist für den Komponenten-Designer erforderlich.
        InitializeComponent()

    End Sub

    'Die Komponente überschreibt den Löschvorgang zum Bereinigen der Komponentenliste.
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

    'Wird vom Komponenten-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Komponenten-Designer erforderlich.
    'Das Bearbeiten ist mit dem Komponenten-Designer möglich.
    'Nehmen Sie keine Änderungen mit dem Code-Editor vor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.VISBOReport = Me.Factory.CreateRibbonTab
        Me.VISBO = Me.Factory.CreateRibbonGroup
        Me.EPReport = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.VISBOReport.SuspendLayout()
        Me.VISBO.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'VISBOReport
        '
        Me.VISBOReport.Groups.Add(Me.VISBO)
        Me.VISBOReport.Label = "VISBO Report"
        Me.VISBOReport.Name = "VISBOReport"
        Me.VISBOReport.Position = Me.Factory.RibbonPosition.BeforeOfficeId("VISBOReport")
        '
        'VISBO
        '
        Me.VISBO.Items.Add(Me.EPReport)
        Me.VISBO.Label = "VISBO"
        Me.VISBO.Name = "VISBO"
        '
        'EPReport
        '
        Me.EPReport.Label = "Einzelprojekt Report"
        Me.EPReport.Name = "EPReport"
        '
        'VisboReportRibbon
        '
        Me.Name = "VisboReportRibbon"
        Me.RibbonType = "Microsoft.Project.Project"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.VISBOReport)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.VISBOReport.ResumeLayout(False)
        Me.VISBOReport.PerformLayout()
        Me.VISBO.ResumeLayout(False)
        Me.VISBO.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents VISBOReport As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents VISBO As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents EPReport As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property VisboReportRibbon() As VisboReportRibbon
        Get
            Return Me.GetRibbon(Of VisboReportRibbon)()
        End Get
    End Property
End Class
