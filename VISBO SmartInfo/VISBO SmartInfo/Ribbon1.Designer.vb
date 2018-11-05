Partial Class Ribbon1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Tab2 = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btnUpdate = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.btnStart = Me.Factory.CreateRibbonButton
        Me.btnFastBack = Me.Factory.CreateRibbonButton
        Me.btnDate = Me.Factory.CreateRibbonButton
        Me.btnFastForward = Me.Factory.CreateRibbonButton
        Me.btnEnd2 = Me.Factory.CreateRibbonButton
        Me.btnToggle = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.btnShowChanges = Me.Factory.CreateRibbonButton
        Me.activateSearch = Me.Factory.CreateRibbonButton
        Me.activateInfo = Me.Factory.CreateRibbonButton
        Me.activateTab = Me.Factory.CreateRibbonButton
        Me.btnFreeze = Me.Factory.CreateRibbonButton
        Me.SmartInfo = Me.Factory.CreateRibbonGroup
        Me.settingsTab = Me.Factory.CreateRibbonButton
        Me.varianten_Tab = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.SmartInfo.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Tab2
        '
        Me.Tab2.Groups.Add(Me.Group2)
        Me.Tab2.Groups.Add(Me.Group3)
        Me.Tab2.Groups.Add(Me.Group4)
        Me.Tab2.Groups.Add(Me.SmartInfo)
        Me.Tab2.Label = "VISBO"
        Me.Tab2.Name = "Tab2"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnUpdate)
        Me.Group2.Label = "Update"
        Me.Group2.Name = "Group2"
        '
        'btnUpdate
        '
        Me.btnUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnUpdate.Image = Global.VISBO_SmartInfo.My.Resources.Resources.refresh
        Me.btnUpdate.Label = "Current"
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.ScreenTip = "VISBO Time-Machine: synchronize with latest version"
        Me.btnUpdate.ShowImage = True
        Me.btnUpdate.SuperTip = "all plan-elements of the VISBO report such as milestones, tasks, VISBO charts and" &
    " tables  are synchronized with the the latest version in the database; "
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnStart)
        Me.Group3.Items.Add(Me.btnFastBack)
        Me.Group3.Items.Add(Me.btnDate)
        Me.Group3.Items.Add(Me.btnFastForward)
        Me.Group3.Items.Add(Me.btnEnd2)
        Me.Group3.Items.Add(Me.btnToggle)
        Me.Group3.Label = "Time Machine"
        Me.Group3.Name = "Group3"
        '
        'btnStart
        '
        Me.btnStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnStart.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_beginning2
        Me.btnStart.Label = "First"
        Me.btnStart.Name = "btnStart"
        Me.btnStart.ScreenTip = "VISBO Time-Machine: synchronize with first version"
        Me.btnStart.ShowImage = True
        Me.btnStart.SuperTip = "all plan-elements of the VISBO report such as milestones, tasks, VISBO charts and" &
    " tables  are synchronized with the the first version in the database; "
        '
        'btnFastBack
        '
        Me.btnFastBack.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnFastBack.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_left1
        Me.btnFastBack.Label = "Backward"
        Me.btnFastBack.Name = "btnFastBack"
        Me.btnFastBack.ScreenTip = "VISBO Time-Machine: synchronize with version 1 month earlier"
        Me.btnFastBack.ShowImage = True
        Me.btnFastBack.SuperTip = "all plan-elements of the VISBO report such as milestones, tasks and VISBO charts " &
    "and tables  are synchronized with the version 1 month earlier in the database; "
        '
        'btnDate
        '
        Me.btnDate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnDate.Image = Global.VISBO_SmartInfo.My.Resources.Resources.calendar
        Me.btnDate.Label = "Date"
        Me.btnDate.Name = "btnDate"
        Me.btnDate.ShowImage = True
        Me.btnDate.SuperTip = "all plan-elements of the VISBO report such as milestones, tasks, VISBO charts and" &
    " tables  are synchronized with the version of the selected date and time in the " &
    "database"
        '
        'btnFastForward
        '
        Me.btnFastForward.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnFastForward.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_right1
        Me.btnFastForward.Label = "Foreward"
        Me.btnFastForward.Name = "btnFastForward"
        Me.btnFastForward.ScreenTip = "VISBO Time-Machine: synchronize with version 1 month later"
        Me.btnFastForward.ShowImage = True
        Me.btnFastForward.SuperTip = "all plan-elements of the VISBO report such as milestones, tasks and VISBO charts " &
    "and tables  are synchronized with the the version 1 month later in the database;" &
    " "
        '
        'btnEnd2
        '
        Me.btnEnd2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnEnd2.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_end1
        Me.btnEnd2.Label = "Last"
        Me.btnEnd2.Name = "btnEnd2"
        Me.btnEnd2.ScreenTip = "VISBO Time-Machine: synchronize with latest version"
        Me.btnEnd2.ShowImage = True
        Me.btnEnd2.SuperTip = "all plan-elements of the VISBO report such as milestones, tasks and VISBO charts " &
    "and tables  are synchronized with the latest version in the database; "
        '
        'btnToggle
        '
        Me.btnToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnToggle.Image = Global.VISBO_SmartInfo.My.Resources.Resources.undo
        Me.btnToggle.Label = "Toggle"
        Me.btnToggle.Name = "btnToggle"
        Me.btnToggle.ShowImage = True
        Me.btnToggle.SuperTip = "all plan-elements of the VISBO report such as milestones, tasks, VISBO charts and" &
    " tables  are synchronized. With this button you may toggle two versions in the d" &
    "atabase"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.btnShowChanges)
        Me.Group4.Items.Add(Me.activateSearch)
        Me.Group4.Items.Add(Me.activateInfo)
        Me.Group4.Items.Add(Me.activateTab)
        Me.Group4.Items.Add(Me.btnFreeze)
        Me.Group4.Label = "Actions"
        Me.Group4.Name = "Group4"
        '
        'btnShowChanges
        '
        Me.btnShowChanges.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnShowChanges.Image = Global.VISBO_SmartInfo.My.Resources.Resources.shape_triangle
        Me.btnShowChanges.Label = "Difference"
        Me.btnShowChanges.Name = "btnShowChanges"
        Me.btnShowChanges.ScreenTip = "VISBO Time-Machine: show differences"
        Me.btnShowChanges.ShowImage = True
        Me.btnShowChanges.SuperTip = "shows differences in milestone and task dates between current version and last ac" &
    "tive version in the database"
        '
        'activateSearch
        '
        Me.activateSearch.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.activateSearch.Image = Global.VISBO_SmartInfo.My.Resources.Resources.find1
        Me.activateSearch.Label = "Search"
        Me.activateSearch.Name = "activateSearch"
        Me.activateSearch.ScreenTip = "show Search Pane for various information categories"
        Me.activateSearch.ShowImage = True
        Me.activateSearch.SuperTip = "user may search for different categories such as certain name, all elements with " &
    "certain traffic light, elements where a certain person is responsible, etc. "
        '
        'activateInfo
        '
        Me.activateInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.activateInfo.Image = Global.VISBO_SmartInfo.My.Resources.Resources.about
        Me.activateInfo.Label = "Properties"
        Me.activateInfo.Name = "activateInfo"
        Me.activateInfo.ScreenTip = "show Properties Pane of 1 selected plan-element"
        Me.activateInfo.ShowImage = True
        Me.activateInfo.SuperTip = "shows for the selected milestone or phase properties such as name, traffic light " &
    "and according explanation, deliverables and person responsible for the milestone" &
    " / tasks"
        '
        'activateTab
        '
        Me.activateTab.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.activateTab.Image = Global.VISBO_SmartInfo.My.Resources.Resources.pen_blue1
        Me.activateTab.Label = "Annotate"
        Me.activateTab.Name = "activateTab"
        Me.activateTab.ScreenTip = "annotate milestones and phases with name and/or date "
        Me.activateTab.ShowImage = True
        Me.activateTab.SuperTip = "annotates all selected milestones and tasks with name and/or date"
        '
        'btnFreeze
        '
        Me.btnFreeze.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnFreeze.Image = Global.VISBO_SmartInfo.My.Resources.Resources.snowflake
        Me.btnFreeze.Label = "Freeze/Defreeze"
        Me.btnFreeze.Name = "btnFreeze"
        Me.btnFreeze.ShowImage = True
        Me.btnFreeze.SuperTip = resources.GetString("btnFreeze.SuperTip")
        '
        'SmartInfo
        '
        Me.SmartInfo.Items.Add(Me.settingsTab)
        Me.SmartInfo.Items.Add(Me.varianten_Tab)
        Me.SmartInfo.Items.Add(Me.Button1)
        Me.SmartInfo.Name = "SmartInfo"
        '
        'settingsTab
        '
        Me.settingsTab.Label = "Settings"
        Me.settingsTab.Name = "settingsTab"
        '
        'varianten_Tab
        '
        Me.varianten_Tab.Label = "Variants"
        Me.varianten_Tab.Name = "varianten_Tab"
        Me.varianten_Tab.Visible = False
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Label = "Smartify Chart"
        Me.Button1.Name = "Button1"
        Me.Button1.ScreenTip = "embeds information about data source in Chart and uploads timestamped data source" &
    ""
        Me.Button1.Visible = False
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Tab2)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Tab2.ResumeLayout(False)
        Me.Tab2.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.SmartInfo.ResumeLayout(False)
        Me.SmartInfo.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnUpdate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnStart As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFastBack As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnShowChanges As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SmartInfo As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents settingsTab As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents varianten_Tab As Microsoft.Office.Tools.Ribbon.RibbonButton

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub Ribbon1_Close(sender As Object, e As EventArgs) Handles Me.Close

        Try
            My.Settings.Save()
        Catch ex As Exception

        End Try

        Try
            Call closeExcelAPP()
        Catch ex As Exception

        End Try

    End Sub

    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFastForward As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnEnd2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents activateSearch As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents activateInfo As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents activateTab As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFreeze As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnToggle As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
