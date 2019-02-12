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
    Private _portfolioAggregationRoleIDs() As Integer
    ' wird benötigt, um bestimmen zu können, welche projectboard Funktionalität erlaubt / nicht erlaubt ist 
    Private _nonAllowance() As String

    Public Sub New()
        _userName = ""
        _userID = ""
        _customUserRole = ptCustomUserRoles.Alles
        _specifics = Nothing
        '_portfolioAggregationRoles = {""}
        _portfolioAggregationRoleIDs = {1}
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
                _nonAllowance = {"Pt5G2B1", "Pt5G2B4", "PT5G3M",
                                 "PT4G1M1-2", "PT4G1M1-3", "PT4G1B14", "PT4G2M",
                                 "PTneu", "PTedit", "PTview",
                                 "PTfilter", "PTsort", "PT0G1s9",
                                 "PTOPTB1", "PTreport",
                                 "PTeinst", "PThelp", "PTWebServer"}

            Case ptCustomUserRoles.PortfolioManager
                '_nonAllowance = {"PT4G1M1-1", "PT4G1M1-2",
                '                 "PTview", "PTfilter", "PTWebServer"}
                _nonAllowance = {"PT4G1M1-1", "PT4G1M1-2",
                                 "PTview", "PTWebServer"}

            Case ptCustomUserRoles.ProjektLeitung
                _nonAllowance = {"Pt5G2B1", "Pt5G2B4", "Pt5G3B1", "PT4G1M1-1",
                                 "PT2G1B1", "PT2G1B3", "PTfilter", "PTsort", "PTeinst", "PThelp",
                                 "PTWebServer"}


            Case ptCustomUserRoles.RessourceManager

                _nonAllowance = {"Pt5G2B1", "Pt5G2B4", "PT5G3M", "Pt5G3B1",
                                 "PT4G1B8", "PT4G1B12", "PT4G1B11",
                                 "PT4G1M1-2", "PT4G1M1-3",
                                 "PT2G1M2B3", "PT2G1M2B8",
                                 "PT4G1M1B2", "PT2G1B1", "PT2G1B3",
                                 "PTfreezeB1", "PTfreezeB2", "PT2G1M1B4", "PT2G1split",
                                 "PTview", "PTsort", "PTeinst", "PThelp",
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
    Public Function isAllowedToSee(ByVal nameOrID As String) As Boolean
        Dim isAllowed As Boolean = False

        If nameOrID = "" Then

            isAllowed = (myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager) Or
                        (myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung) Or
                        (myCustomUserRole.customUserRole = ptCustomUserRoles.Alles)
        Else
            Dim teamID As Integer
            Dim roleID As Integer = RoleDefinitions.parseRoleNameID(nameOrID, teamID)
            Dim curRoleDef As clsRollenDefinition = RoleDefinitions.getRoleDefByID(roleID)

            If Not IsNothing(curRoleDef) Then

                Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(roleID, teamID)

                If _customUserRole = ptCustomUserRoles.RessourceManager Then
                    Dim prntTeamID As Integer = -1
                    Dim parentRoleID As Integer = RoleDefinitions.getRoleDefByIDKennung(_specifics, prntTeamID).UID
                    isAllowed = RoleDefinitions.hasAnyChildParentRelationsship(roleNameID, parentRoleID)

                    ' mit dem Folgenden wird sichergestellt, dass ein Ressourcen-Manager , z.B KB1, auch eine Person von KB1 in seiner Eigenschaft als Team-Member sehen kann  
                    If Not isAllowed And teamID > 0 Then
                        Dim roleNameIDBasic As String = RoleDefinitions.bestimmeRoleNameID(roleID, -1)
                        isAllowed = RoleDefinitions.hasAnyChildParentRelationsship(roleNameIDBasic, parentRoleID)
                    End If


                ElseIf _customUserRole = ptCustomUserRoles.PortfolioManager Then
                    isAllowed = _portfolioAggregationRoleIDs.Contains(roleID)
                    If Not isAllowed Then
                        isAllowed = Not RoleDefinitions.hasAnyChildParentRelationsship(roleNameID, _portfolioAggregationRoleIDs)
                    End If

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
            If _customUserRole = ptCustomUserRoles.RessourceManager Then
                encryptedUserRole = visboCrypto.EncryptData(CInt(_customUserRole).ToString & "#" & _specifics)
            Else
                encryptedUserRole = visboCrypto.EncryptData(CInt(_customUserRole).ToString & "#" & "XYZ")
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
        _customUserRole = CType(tmpstr(0), ptCustomUserRoles)
        If _customUserRole = ptCustomUserRoles.RessourceManager Then
            _specifics = CStr(tmpstr(1))
        End If


    End Sub

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
                If [Enum].IsDefined(GetType(ptCustomUserRoles), value) Then
                    _customUserRole = value

                    ' Sonderbehandlung , wenn Portfolio Manager
                    If Not IsNothing(_specifics) And _customUserRole = ptCustomUserRoles.PortfolioManager Then
                        Call setAggregationRoles(_specifics)
                    End If

                End If
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

                ' Sonderbehandlung: wenn es sich um einen Portfolio Manager handelt 
                If _customUserRole = ptCustomUserRoles.PortfolioManager Then
                    Call setAggregationRoles(_specifics)
                End If

            Else
                _specifics = ""
            End If

        End Set
    End Property

    '''' <summary>
    '''' gibt die Namen der AggregationRoleNames in einem String-Array zurück
    '''' hat nur Bedeutung wenn userRole = portfolioMgr
    '''' </summary>
    '''' <returns></returns>
    'Public ReadOnly Property getAggregationRoleNames As String()
    '    Get
    '        getAggregationRoleNames = _portfolioAggregationRoles
    '    End Get
    'End Property

    ''' <summary>
    ''' gibt die Namen der AggregationRoleIDs in einem Integer-Array zurück
    ''' hat nur Bedeutung wenn userRole = portfolioMgr
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getAggregationRoleIDs As Integer()
        Get
            getAggregationRoleIDs = _portfolioAggregationRoleIDs
        End Get
    End Property

    Private Sub setAggregationRoles(ByVal aggregationRoleStr As String)

        Dim tmpStr() As String = aggregationRoleStr.Split(New Char() {CChar(";")})
        Dim tmpCollection As New Collection
        Dim teamID As Integer = -1
        Dim roleUID As String

        For Each tmpName As String In tmpStr

            ' kann eine Uid.string sein, uid;teamID-String oder aber ein rollen-Name 
            If RoleDefinitions.containsNameID(tmpName.Trim) Then
                roleUID = RoleDefinitions.getRoleDefByIDKennung(tmpName.Trim, teamID).UID.ToString
                tmpCollection.Add(roleUID)
            End If


        Next

        If tmpCollection.Count > 0 Then

            'ReDim _portfolioAggregationRoles(tmpCollection.Count - 1)
            ReDim _portfolioAggregationRoleIDs(tmpCollection.Count - 1)

            Dim i As Integer = 0
            ' in tmpCollection sind jetzt ausschließlich RoleNameIDs enthalten
            For Each tmpNameID As String In tmpCollection
                'Dim tmpRoleDef As clsRollenDefinition = RoleDefinitions.getRoledef(tmpNameID)
                '_portfolioAggregationRoles(i) = tmpRoleDef.name
                '_portfolioAggregationRoleIDs(i) = tmpRoleDef.UID
                _portfolioAggregationRoleIDs(i) = CInt(tmpNameID)
                i = i + 1
            Next
        Else
            '_portfolioAggregationRoles = {""}
            _portfolioAggregationRoleIDs = {1}
        End If

    End Sub



End Class
