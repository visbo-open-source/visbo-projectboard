Imports System.Windows.Forms

Public Class frmAllocateRessources

    ' die Rollen-ID in der form roleUid;teamID oder roleUid.tostring bzw. costuid.tostring 
    Public rcNameID As String

    ' die PhaseNameID der Zeile  
    Public phaseNameID As String

    ' das in der Zeile aktive Projekt
    Public hproj As clsProjekt

    Public newValueForRCName As Double

    Public roleSkillValuesToAdd As New SortedList(Of String, Double)

    ' holds the initial sum and sum when being changed ... 
    Private amountToSubstitute As Double

    ' holds the last value a Amount cell contained 
    Private lastValue As Double = 0.0

    ' holds all people alreay in project team
    Private teamList As New Collection

    ' holds all people already allocated in current phase 
    Private teamPhaseList As New Collection

    Private cPhase As clsPhase = Nothing
    Private myRole As clsRolle = Nothing
    Private mySkillName As String = ""

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub

    Private Sub frmAllocateRessources_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If frmCoord(PTfrm.rolecostME, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.top))
            Me.Left = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.left))
        Else
            Me.Top = 60
            Me.Left = 100
        End If

        Dim headerText As String = ""
        If awinSettings.englishLanguage Then
            headerText = "Select People"
        Else
            headerText = "Wählen Sie die Personen"
        End If

        Dim errMsg As String = ""
        Call buildAllocationContent(errMsg)

        If errMsg <> "" Then
            Call MsgBox(errMsg)
        End If

    End Sub

    ''' <summary>
    ''' gets all data from datagridview and pits it into roleSkillValuesToAdd and newValueForRCname
    ''' </summary>
    Private Sub pickupInput()

        If amountToSubstitute >= 0 Then
            newValueForRCName = amountToSubstitute
        Else
            newValueForRCName = 0
        End If

        ' now get all input where values are edited 
        For Each tmpRow As DataGridViewRow In candidatesTable.Rows
            Dim myNewValue As Double = CDbl(tmpRow.Cells.Item(3).Value)
            If myNewValue > 0 Then
                Dim myRoleName As String = CStr(tmpRow.Cells.Item(0).Value)
                Dim myRoleSkillID As String = RoleDefinitions.bestimmeRoleNameID(myRoleName, mySkillName)

                If myRoleSkillID <> "" Then
                    roleSkillValuesToAdd.Add(myRoleSkillID, myNewValue)
                End If
            End If
        Next

    End Sub

    Private Sub buildAllocationContent(ByRef errMsg As String)

        Dim skillID As Integer = -1
        Dim myRoleDef As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(rcNameID, skillID)
        Dim mySkillDef As clsRollenDefinition = Nothing
        If skillID > 0 Then
            mySkillDef = RoleDefinitions.getRoleDefByID(skillID)
            If Not IsNothing(mySkillDef) Then
                mySkillName = mySkillDef.name
            End If
        End If


        cPhase = hproj.getPhaseByID(phaseNameID)
        myRole = cPhase.getRoleByRoleNameID(rcNameID)

        teamList = hproj.getRoleNames

        If Not IsNothing(cPhase) And Not IsNothing(myRole) Then

            teamPhaseList = cPhase.getRoleNames

            If cPhase.hasForecastMonths Then

                Dim foreCastOffset As Integer = 0
                If cPhase.hasActualData Then
                    foreCastOffset = getColumnOfDate(hproj.actualDataUntil) - getColumnOfDate(cPhase.getStartDate) + 1
                End If

                amountToSubstitute = myRole.Xwerte.Sum
                If foreCastOffset > 0 Then
                    ' sum ist calculated from index + 1 
                    amountToSubstitute = myRole.getSumFrom(foreCastOffset - 1)
                End If

                lblOrgaUnitSkill.Text = myRoleDef.name
                If Not IsNothing(mySkillDef) Then
                    lblOrgaUnitSkill.Text = lblOrgaUnitSkill.Text & ", " & mySkillDef.name
                End If

                lblSum.Text = amountToSubstitute.ToString("###0.#")

                Dim candidatesList As SortedList(Of Double, Integer) = cPhase.getCandidates(rcNameID, 1, amountToSubstitute)

                Dim tableIndex As Integer = 0
                For Each kvp As KeyValuePair(Of Double, Integer) In candidatesList
                    Dim curRoleDef As clsRollenDefinition = RoleDefinitions.getRoleDefByID(kvp.Value)
                    Dim candidatesName As String = curRoleDef.name
                    Dim freeCapacity As Double = System.Math.Truncate(10 * kvp.Key) / 10

                    Dim values As Object()
                    ReDim values(3)
                    values(0) = candidatesName
                    values(1) = freeCapacity
                    values(2) = " "
                    If curRoleDef.isExternRole Then
                        values(2) = "Yes"
                    End If
                    values(3) = 0

                    With candidatesTable
                        .Rows.Insert(tableIndex, values)
                    End With
                Next

            Else
                ' phase does not have any forecast months, so it is not possible to change anything anymore
                lblOrgaUnitSkill.Text = "Phase with no forecast months anymore - so changing is not possible"
                lblSum.Text = ""
            End If
        End If

        If candidatesTable.Rows.Count > 0 Then
            candidatesTable.Rows.Item(0).Selected = True
            candidatesTable.Rows.Item(0).Cells(3).Selected = True
        End If

    End Sub



    ''' <summary>
    ''' called when Edit ends
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub candidatesTable_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles candidatesTable.CellEndEdit
        Dim myCol As Integer = e.ColumnIndex


        If myCol = 3 Then
            Dim currentValue As Double = CDbl(candidatesTable.CurrentCell.Value)
            Dim difference As Double = currentValue - lastValue

            If difference <> 0 Then
                ' Adjust amountToSubstitute
                amountToSubstitute = amountToSubstitute - difference
                If amountToSubstitute < 0 Then
                    amountToSubstitute = 0
                End If
                If amountToSubstitute > myRole.Xwerte.Sum Then
                    amountToSubstitute = myRole.Xwerte.Sum
                End If
                lblSum.Text = amountToSubstitute.ToString("###0.#")
            End If

        End If
    End Sub

    Private Sub candidatesTable_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles candidatesTable.CellEnter
        Dim myCol As Integer = e.ColumnIndex
        If myCol = 3 Then
            lastValue = CDbl(candidatesTable.CurrentCell.Value)
        End If
    End Sub

    Private Sub candidatesTable_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles candidatesTable.CellFormatting
        If e.ColumnIndex = 0 AndAlso e.Value IsNot Nothing Then
            If teamPhaseList.Contains(CStr(e.Value)) Then
                e.CellStyle.BackColor = Drawing.Color.Azure
            ElseIf teamList.Contains(CStr(e.Value)) Then
                e.CellStyle.BackColor = Drawing.Color.Azure
            End If
        End If
    End Sub

    Private Sub candidatesTable_CellToolTipTextNeeded(sender As Object, e As DataGridViewCellToolTipTextNeededEventArgs) Handles candidatesTable.CellToolTipTextNeeded

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub okBtn_Click(sender As Object, e As EventArgs) Handles okBtn.Click
        Call pickupInput()
    End Sub
End Class