''' <summary>
''' tk 6.12.18 wird in den Role-TreeViews benötigt,um Information an TreeNode zu hängen, in Form von Tags
''' </summary>
Public Class clsNodeRoleTag
    Public Property pTag As Char
    Public Property isRole As Boolean
    Public Property isSkill As Boolean
    Public Property isTeamMember As Boolean
    Public Property membershipID As Integer
    Public Property membershipPrz As Double

    Public Sub New()
        _pTag = CChar("")
        _isRole = True
        _isSkill = False
        _isTeamMember = False
        _membershipID = -1
        _membershipPrz = 0.0
    End Sub
End Class
