Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Erforderlich für die Unterstützung des Windows.Forms-Klassenkompositions-Designers
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Dieser Aufruf ist für den Komponenten-Designer erforderlich.
        InitializeComponent()

    End Sub

    'Die Komponente überschreibt den Löschvorgang zum Bereinigen der Komponentenliste.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.oneClickPPT = Me.Factory.CreateRibbonTab
        Me.VISBO = Me.Factory.CreateRibbonGroup
        Me.EinzelprojektReport = Me.Factory.CreateRibbonButton
        Me.DBspeichern = Me.Factory.CreateRibbonButton
        Me.Baseline = Me.Factory.CreateRibbonButton
        Me.Einstellungen = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.oneClickPPT.SuspendLayout()
        Me.VISBO.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'oneClickPPT
        '
        Me.oneClickPPT.Groups.Add(Me.VISBO)
        Me.oneClickPPT.Label = "VISBO"
        Me.oneClickPPT.Name = "oneClickPPT"
        '
        'VISBO
        '
        Me.VISBO.Items.Add(Me.EinzelprojektReport)
        Me.VISBO.Items.Add(Me.DBspeichern)
        Me.VISBO.Items.Add(Me.Baseline)
        Me.VISBO.Items.Add(Me.Einstellungen)
        Me.VISBO.Name = "VISBO"
        '
        'EinzelprojektReport
        '
        Me.EinzelprojektReport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.EinzelprojektReport.Image = CType(resources.GetObject("EinzelprojektReport.Image"), System.Drawing.Image)
        Me.EinzelprojektReport.Label = "Einzelprojekt Report"
        Me.EinzelprojektReport.Name = "EinzelprojektReport"
        Me.EinzelprojektReport.ShowImage = True
        Me.EinzelprojektReport.Visible = False
        '
        'DBspeichern
        '
        Me.DBspeichern.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.DBspeichern.Image = CType(resources.GetObject("DBspeichern.Image"), System.Drawing.Image)
        Me.DBspeichern.Label = "Publish in VISBO"
        Me.DBspeichern.Name = "DBspeichern"
        Me.DBspeichern.ShowImage = True
        '
        'Baseline
        '
        Me.Baseline.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Baseline.Image = CType(resources.GetObject("Baseline.Image"), System.Drawing.Image)
        Me.Baseline.Label = "Publish Baseline in VISBO"
        Me.Baseline.Name = "Baseline"
        Me.Baseline.ScreenTip = "Publish Baseline in VISBO"
        Me.Baseline.ShowImage = True
        Me.Baseline.Visible = False
        '
        'Einstellungen
        '
        Me.Einstellungen.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Einstellungen.Image = CType(resources.GetObject("Einstellungen.Image"), System.Drawing.Image)
        Me.Einstellungen.Label = "Einstellungen"
        Me.Einstellungen.Name = "Einstellungen"
        Me.Einstellungen.ScreenTip = "Einstellung"
        Me.Einstellungen.ShowImage = True
        Me.Einstellungen.Visible = False
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Project.Project"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.oneClickPPT)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.oneClickPPT.ResumeLayout(False)
        Me.oneClickPPT.PerformLayout()
        Me.VISBO.ResumeLayout(False)
        Me.VISBO.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents oneClickPPT As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents VISBO As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents EinzelprojektReport As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Baseline As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents DBspeichern As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Einstellungen As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
