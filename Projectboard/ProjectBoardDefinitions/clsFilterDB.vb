Public Class clsFilterDB

    Public filterPhase As List(Of String)
    Public filterMilestone As List(Of String)
    Public filterRolle As List(Of String)
    Public filterCost As List(Of String)
    Public filterTyp As List(Of String)
    Public filterBU As List(Of String)
    Public name As String
    Public selFilter As Boolean
    Public Id As String



    Sub copyfrom(ByRef item As clsFilter, ByVal sf As Boolean)

        Me.name = item.name
        Me.selFilter = sf
        For Each c In item.Phases
            Me.filterPhase.Add(CStr(c))
        Next
        For Each c In item.Milestones
            Me.filterMilestone.Add(CStr(c))
        Next
        For Each c In item.Roles
            Me.filterRolle.Add(CStr(c))
        Next
        For Each c In item.Costs
            Me.filterCost.Add(CStr(c))
        Next
        For Each c In item.Typs
            Me.filterTyp.Add(CStr(c))
        Next
        For Each c In item.BUs
            Me.filterBU.Add(CStr(c))
        Next


    End Sub

    Sub copyto(ByRef item As clsFilter)

        item.name = Me.name
        For Each c In Me.filterPhase
            item.Phases.Add(c, c)
        Next
        For Each c In Me.filterMilestone
            item.Milestones.Add(c, c)
        Next
        For Each c In Me.filterRolle
            item.Roles.Add(c, c)
        Next
        For Each c In Me.filterCost
            item.Costs.Add(c, c)
        Next
        For Each c In Me.filterTyp
            item.Typs.Add(c, c)
        Next
        For Each c In Me.filterBU
            item.BUs.Add(c, c)
        Next

    End Sub

    Sub New()
        filterPhase = New List(Of String)
        filterMilestone = New List(Of String)
        filterRolle = New List(Of String)
        filterCost = New List(Of String)
        filterTyp = New List(Of String)
        filterBU = New List(Of String)
    End Sub

End Class
