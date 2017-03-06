Imports MongoDB.Bson
''' <summary>
''' die Datenbank Klasse für Projekt-Write Protections
''' </summary>
''' <remarks></remarks>
Public Class clsWriteProtectionItemDB
    Public Id As ObjectId
    Public pName As String
    Public vName As String
    Public type As Integer
    Public userName As String
    Public isProtected As Boolean
    Public permanent As Boolean
    Public lastDateSet As Date
    Public lastDateReleased As Date

    Public Sub copyFrom(ByVal wpItem As clsWriteProtectionItem)

        With wpItem

            Me.pName = getPnameFromKey(.pvName)
            Me.vName = getVariantnameFromKey(.pvName)
            Me.type = .type
            Me.userName = .userName
            Me.isProtected = .isProtected
            Me.permanent = .permanent
            Me.lastDateSet = .lastDateSet.ToUniversalTime
            Me.lastDateReleased = .lastDateReleased.ToUniversalTime

        End With
    End Sub

    Public Sub copyTo(ByRef wpItem As clsWriteProtectionItem)

        With wpItem

            .pvName = calcProjektKey(Me.pName, Me.vName)
            .type = Me.type
            .userName = Me.userName
            .isProtected = Me.isProtected
            .permanent = Me.permanent
            .lastDateSet = Me.lastDateSet.ToLocalTime
            .lastDateReleased = Me.lastDateSet.ToLocalTime

        End With
    End Sub

    Sub New()

    End Sub

End Class
