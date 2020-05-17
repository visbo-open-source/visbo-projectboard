''' <summary>
''' siehe https://visbogmbh.atlassian.net/wiki/spaces/VS/pages/231735299/Erweiterungen+in+Datenmodell+getriggert+durch+Allianz
''' 
''' </summary>
Public Class clsCustomUserRole

    Private _userName As String
    Private _userID As String
    Private _customUserRole As ptCustomUserRoles
    ' gibt im Falle resource Mgr an, welche Orga-Einhairt er nur sehen darf 
    Private _specifics As String
    ' gibt im Falle Portfolio Manager an, welche Rollen ggf aggregiert werden sollen 
    'Private _portfolioAggregationRoles() As String

    Private _nonAllowance() As String

    Public Sub New()
        _userName = ""
        _userID = ""
        _customUserRole = ptCustomUserRoles.OrgaAdmin
        _specifics = ""
        _nonAllowance = {""}
    End Sub

    ''' <summary>
    ''' setzt , in Abhängigkeit von _customUserRole die Menu-Punkt Allowance
    ''' muss aufgerufen werden, sobald eine customUSerRole gewählt wurde 
    ''' </summary>
    Public Sub setNonAllowances()

        Select Case _customUserRole

            Case ptCustomUserRoles.Alles
                _nonAllowance = {""}

            Case ptCustomUserRoles.OrgaAdmin
                _nonAllowance = {"Pt5G2B1", "Pt5G2B4", "PT5G3M", "PT2G1split",
                                 "PT4G1M1-2", "PT4G1M1-3", "PT4G2M-1", "PT2G1M2B2",
                                 "PT2G2B2", "separator3", "PTfreezeB1", "PTfreezeB2",
                                 "PTview", "PTmassEdit",
                                 "PTfilter", "PTsort", "PT0G1s9",
                                 "PTOPTB1",
                                 "PThelp", "PTWebServer"}

            Case ptCustomUserRoles.PortfolioManager
                '_nonAllowance = {"PT4G1M1-1", "PT4G1M1-2",
                '                 "PTview", "PTfilter", "PTWebServer"}
                _nonAllowance = {"PT4G1M1-1", "PT4G1B12", "PT4G1B15", "PT4G1B16", "PT4G1M0B2", "PTfilter",
                                 "PT2G1M2B8", "PT2G1B1", "PT2G1M1B4",
                                 "PT0G1B3", "PT7G1M2", "PTXG1B3", "PTXG1B8", "PT1G1B6",
                                 "PTWebServer", "PThelp"}

            Case ptCustomUserRoles.ProjektLeitung
                _nonAllowance = {"Pt5G2B1", "Pt5G2B4", "Pt5G3B1", "PT4G1M1-1", "PT4G1B12", "PT4G1B15", "PT4G1B16",
                                 "PT2G1B1", "PT2G1B3", "PT2G1M2B3", "PTfilter", "PTsort", "PThelp",
                                 "PT2G1B1", "PT2G1M1B4",
                                 "PTWebServer"}


            Case ptCustomUserRoles.RessourceManager

                _nonAllowance = {"Pt5G2B1", "Pt5G2B4", "PT5G3M", "Pt5G3B1",
                                 "PT2G1B1", "PT2G1M1B4",
                                 "PT4G1M1-1", "PT4G1M1-2", "PT4G1M1-3", "PT4G1M0B2", "PT4G1B8", "PT4G1B12", "PT4G1B15", "PT4G1B16", "PT4G1B11",
                                 "PT4G2B3", "PT2G1M2B3", "PT2G1M2B8",
                                 "PT0G1B3", "PT7G1M2", "PTXG1B3", "PTXG1B8",
                                 "PT4G1M1B2", "PT2G1B1", "PT2G1B3",
                                 "PTfreezeB1", "PTfreezeB2", "PT2G1M1B4", "PT2G1split",
                                 "PTview", "PTsort", "PTfilter", "PThelp", "PT1G1B6",
                                 "PTWebServer"}

                ' Team-Manager und Ressourcen-Manager solten die gleichen Funktionen sehen / nicht sehen 
            Case ptCustomUserRoles.TeamManager

                _nonAllowance = {"Pt5G2B1", "Pt5G2B4", "PT5G3M", "Pt5G3B1", "PT2G1B1", "PT2G1M1B4",
                                 "PT4G1M1-1", "PT4G1M1-2", "PT4G1M1-3", "PT4G1M0B2", "PT4G1B8", "PT4G1B15", "PT4G1B16", "PT4G1B12", "PT4G1B11",
                                 "PT4G2B3", "PT2G1M2B3", "PT2G1M2B8",
                                 "PT0G1B3", "PT7G1M2", "PTXG1B3", "PTXG1B8",
                                 "PT4G1M1B2", "PT2G1B1", "PT2G1B3",
                                 "PTfreezeB1", "PTfreezeB2", "PT2G1M1B4", "PT2G1split",
                                 "PTview", "PTsort", "PTfilter", "PThelp", "PT1G1B6",
                                 "PTWebServer"}

                ' internal Viewer sollte die gleichen Funktionen wie Team und Ressourcen Manager sehen, ausser alles was mit editieren zu tun hat 
            Case ptCustomUserRoles.InternalViewer

                _nonAllowance = {"Pt5G2B1", "Pt5G2B4", "PT5G3M", "Pt5G3B1", "PT4G1B12", "PT4G1B15", "PT4G1B16", "PT2G1B1", "PT2G1M1B4",
                                 "PT2G1M2B1", "PT2G1M2B8", "PT2G1M2B2", "PT2G2B5", "PT2G1M1B3",
                                 "PT4G1M",
                                 "PT0G1B3", "PT7G1M2", "PTXG1B3", "PTXG1B8",
                                 "PT5G2split", "PT5G2", "PT5G2M",
                                 "PT4G1M1B2",
                                 "PTneu",
                                 "PTview", "PTsort", "PTfilter", "PThelp", "PT1G1B6",
                                 "PTWebServer"}
            Case Else
                _nonAllowance = {""}
        End Select

    End Sub

    ''' <summary>
    ''' bestimmt in Abhängigkeit von der customUSerRole, ob eine bestimmte Person, Orga-Einheit gesehen werden darf ... 
    ''' roleName darf Name der Rolle oder NameIDstr sein oder "" für 'Alles'
    ''' </summary>
    ''' <param name="nameOrID"></param>
    ''' <returns></returns>
    Public Function isAllowedToSee(ByVal nameOrID As String,
                                   Optional includingVirtualChilds As Boolean = False) As Boolean
        ' die Aufruf-Schnittstelle wurde geändert , includingVirtualChilds ist jetzt immer true 
        ' tk Änderung 18.1 includingVirtualChilds ist immer true 


        Dim isAllowed As Boolean = False

        If nameOrID = "" Then

            isAllowed = (myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager) Or
                        (myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung) Or
                        (myCustomUserRole.customUserRole = ptCustomUserRoles.InternalViewer) Or
                        (myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin) Or
                        (myCustomUserRole.customUserRole = ptCustomUserRoles.Alles)
        Else
            Dim teamID As Integer
            Dim roleID As Integer = RoleDefinitions.parseRoleNameID(nameOrID, teamID)
            Dim curRoleDef As clsRollenDefinition = RoleDefinitions.getRoleDefByID(roleID)


            If Not IsNothing(curRoleDef) Then

                Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(roleID, teamID)

                If _customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
                    Dim prntTeamID As Integer = -1
                    Dim restrictedToRoleDef As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(specifics, prntTeamID)
                    Dim restrictedToRoleID As Integer = restrictedToRoleDef.UID
                    ' tk 18.1.20 bei einem Team kommt nur dann true raus, wenn alle Team-Mitglieder in der restrictedRoleID sind, 
                    ' deshalb wurde das von dem anderen Aufruf ersetzt 
                    ' ALt vor 18.1 20
                    'isAllowed = RoleDefinitions.hasAnyChildParentRelationsship(roleNameID, restrictedToRoleID, includingVirtualChilds:=includingVirtualChilds)
                    ' mit dem Folgenden wird sichergestellt, dass ein Ressourcen-Manager , z.B KB1, auch eine Person von KB1 in seiner Eigenschaft als Team-Member sehen kann
                    ' nicht includingVirtualChilds, weil das eine PErson betrifft ..
                    'If Not isAllowed And teamID > 0 Then
                    '    Dim roleNameIDBasic As String = RoleDefinitions.bestimmeRoleNameID(roleID, -1)
                    '    isAllowed = RoleDefinitions.hasAnyChildParentRelationsship(roleNameIDBasic, restrictedToRoleID)
                    'End If
                    ' Ende Alt vor 18.1.20

                    ' Neu seit 18.1.20
                    If restrictedToRoleDef.isTeam Or restrictedToRoleDef.isTeamParent Then
                        isAllowed = RoleDefinitions.hasAnyChildParentRelationsship(roleNameID, restrictedToRoleID)
                    Else
                        Dim tmpergList As List(Of Integer) = RoleDefinitions.getCommonChildsOfParents(roleID, restrictedToRoleID)
                        isAllowed = tmpergList.Count > 0
                    End If
                    ' Ende Neu seit 18.1.20

                    'ElseIf _customUserRole = ptCustomUserRoles.PortfolioManager Then
                    '    Dim idArray() As Integer = getAggregationRoleIDs()
                    '    If Not IsNothing(idArray) Then
                    '        isAllowed = idArray.Contains(roleID)
                    '        If Not isAllowed Then
                    '            isAllowed = Not RoleDefinitions.hasAnyChildParentRelationsship(roleNameID, idArray)
                    '        End If
                    '    Else
                    '        isAllowed = True
                    '    End If


                ElseIf _customUserRole = ptCustomUserRoles.ProjektLeitung Or
                       _customUserRole = ptCustomUserRoles.InternalViewer Or
                       _customUserRole = ptCustomUserRoles.PortfolioManager Or
                       _customUserRole = ptCustomUserRoles.OrgaAdmin Or
                       _customUserRole = ptCustomUserRoles.Alles Then
                    isAllowed = True
                Else
                    isAllowed = False
                End If

            End If
        End If



        isAllowedToSee = isAllowed
    End Function

    ''' <summary>
    ''' gibt an, ob die userRole für die angegebene MenuID berechtigt ist, dass heisst nicht in der nonAllowance aufgeführt ist 
    ''' true
    ''' </summary>
    ''' <param name="menuID"></param>
    ''' <returns></returns>
    Public Function isEntitledForMenu(ByVal menuID As String) As Boolean
        isEntitledForMenu = Not _nonAllowance.Contains(menuID)
    End Function

    ''' <summary>
    ''' verschlüsselt die UserRole, dabei wird die Kennziffer customUserRole und specifics verschlüsselt, sofern es sich um 
    ''' eine Ressource-Manager Rolel handelt 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property encrypt() As String
        Get
            Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)
            Dim encryptedUserRole As String = ""
            If _customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
                encryptedUserRole = visboCrypto.EncryptData(CInt(customUserRole).ToString & "#" & specifics)
            Else
                encryptedUserRole = visboCrypto.EncryptData(CInt(customUserRole).ToString & "#" & "XYZ")
            End If

            encrypt = encryptedUserRole
        End Get
    End Property

    ''' <summary>
    ''' setzt in der aktuellen Instanz die customUserRole und, falls RessourceManager, die specifics entsprechend 
    ''' </summary>
    ''' <param name="encryptedText"></param>
    Public Sub decrypt(ByVal encryptedText As String)

        Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)
        Dim decryptedText As String = visboCrypto.DecryptData(encryptedText)
        Dim tmpstr() As String = decryptedText.Split(New Char() {CChar("#")})
        customUserRole = CType(tmpstr(0), ptCustomUserRoles)
        If customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
            specifics = CStr(tmpstr(1))
        End If


    End Sub
    ''' <summary>
    ''' gibt den Namen des Referats bzw. des Internal-Viewers zurück
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property specificsName() As String
        Get
            Dim tmpResult As String = ""
            Dim teamID As Integer = -1
            If _customUserRole = ptCustomUserRoles.RessourceManager Or
                    _customUserRole = ptCustomUserRoles.TeamManager Or
                    customUserRole = ptCustomUserRoles.InternalViewer Then

                tmpResult = RoleDefinitions.getRoleDefByIDKennung(_specifics, teamID).name

            Else
                tmpResult = ""
            End If

            specificsName = tmpResult
        End Get
    End Property

    Public Property userName As String
        Get
            userName = _userName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _userName = value
            End If
        End Set
    End Property

    Public Property userID As String
        Get
            userID = _userID
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _userID = value
            End If
        End Set
    End Property

    Public Property customUserRole As ptCustomUserRoles
        Get
            customUserRole = _customUserRole
        End Get
        Set(value As ptCustomUserRoles)
            If Not IsNothing(value) Then
                _customUserRole = value
            Else
                _customUserRole = ptCustomUserRoles.OrgaAdmin
            End If
        End Set
    End Property

    Public Property specifics As String
        Get
            specifics = _specifics
        End Get
        Set(value As String)

            If Not IsNothing(value) Then
                _specifics = value
            Else
                _specifics = ""
            End If

        End Set
    End Property


    ''' <summary>
    ''' gibt die Namen der AggregationRoleIDs in einem Integer-Array zurück
    ''' Voraussetzung: specifics enthält nur valide IDs
    ''' </summary>
    ''' <returns></returns>
    Public Function getAggregationRoleIDs() As Integer()
        Dim result() As Integer = Nothing

        If specifics <> "" And _customUserRole = ptCustomUserRoles.PortfolioManager Then

            Dim tmpStr() As String = specifics.Split(New Char() {CChar(";")})
            Dim i As Integer = 0
            If tmpStr.Length > 0 Then
                ReDim result(tmpStr.Length - 1)
                For Each tmpName As String In tmpStr
                    result(i) = CInt(tmpName)
                    i = i + 1
                Next
            End If

        End If

        getAggregationRoleIDs = result

    End Function





End Class
