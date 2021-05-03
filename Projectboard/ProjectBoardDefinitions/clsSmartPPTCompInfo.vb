Public Class clsSmartPPTCompInfo


    Public Property hproj As clsProjekt
    'Public Property hproj2 As clsProjekt
    'Public Property vglProj As clsProjekt
    Public Property bigType As ptReportBigTypes

    Public Property detailID As ptReportComponents

    'Public Property chartTyp As PTChartTypen
    'Public Property vergleichsArt As PTVergleichsArt
    'Public Property vergleichsTyp As PTVergleichsTyp
    'Public Property vergleichsDatum As Date
    'Public Property einheit As PTEinheiten
    Public Property elementTyp As ptElementTypen
    Public Property q1 As String
    Public Property q2 As String
    Public Property text As String


    Private _pName As String
    ''' <summary>
    ''' wenn ein hproj bereits angegeben ist, nimmt er immer den Namen des hproj
    ''' </summary>
    ''' <returns></returns>
    Public Property pName As String
        Get
            If Not IsNothing(_hproj) Then
                pName = _hproj.name
            Else
                pName = _pName
            End If
        End Get
        Set(value As String)
            If IsNothing(_hproj) Then
                _pName = value
            End If
        End Set
    End Property
    Private _vName As String
    ''' <summary>
    ''' wenn ein hproj bereits angegeben ist, gibt er immer den Namen des hproj zurück
    ''' </summary>
    ''' <returns></returns>
    Public Property vName As String
        Get
            If Not IsNothing(_hproj) Then
                vName = _hproj.variantName
            Else
                vName = _vName
            End If
        End Get
        Set(value As String)
            If IsNothing(_hproj) Then
                _vName = value
            End If
        End Set
    End Property


    Private _prPF As ptPRPFType
    Public Property prPF As ptPRPFType
        Get
            If Not IsNothing(_hproj) Then
                prPF = CType(hproj.projectType, ptPRPFType)
            Else
                prPF = _prPF
            End If
        End Get
        Set(value As ptPRPFType)
            If IsNothing(_hproj) Then
                _prPF = value
            End If
        End Set
    End Property

    Private _vpid As String
    Public Property vpid As String
        Get
            vpid = _vpid
        End Get
        Set(value As String)
            _vpid = value
        End Set
    End Property

    Private _zeitRaumLeft As Date
    Public Property zeitRaumLeft As Date
        Get
            If Not IsNothing(_zeitRaumLeft) Then
                zeitRaumLeft = _zeitRaumLeft
            Else
                zeitRaumLeft = Date.MinValue
            End If
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                If value >= StartofCalendar Then
                    _zeitRaumLeft = value
                Else
                    _zeitRaumLeft = StartofCalendar
                End If
            End If

        End Set
    End Property

    Private _zeitRaumRight As Date

    Public Sub New()

    End Sub

    Public Property zeitRaumRight As Date
        Get
            If Not IsNothing(_zeitRaumLeft) Then
                zeitRaumRight = _zeitRaumRight
            Else
                zeitRaumRight = Date.MinValue
            End If
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                If value >= StartofCalendar Then
                    _zeitRaumRight = value
                Else
                    _zeitRaumRight = StartofCalendar
                End If
            End If

        End Set
    End Property

    Public ReadOnly Property hasValidZeitraum() As Boolean
        Get
            If Not (IsNothing(_zeitRaumLeft) Or IsNothing(_zeitRaumRight)) Then
                hasValidZeitraum = ((getColumnOfDate(_zeitRaumRight) > getColumnOfDate(_zeitRaumLeft)) And (_zeitRaumLeft >= StartofCalendar))
            Else
                hasValidZeitraum = False
            End If

        End Get
    End Property


    Public ReadOnly Property q1Bezeichner As String
        Get
            Dim tmpResult As String = ""

            If _q1 = "" Then
                tmpResult = ""
            Else
                If Not IsNothing(RoleDefinitions) Then
                    Dim tmpTeamID As Integer = -1
                    Dim tmpRoleID As Integer = RoleDefinitions.parseRoleNameID(_q1, tmpTeamID)
                    If tmpRoleID > 0 Then
                        tmpResult = RoleDefinitions.getRoleDefByID(tmpRoleID).name
                    End If
                End If
            End If
            q1Bezeichner = tmpResult
        End Get
    End Property

    Public ReadOnly Property q2Bezeichner As String
        Get
            Dim tmpResult As String = ""

            If _q2 = "" Then
                tmpResult = ""
            Else
                If Not IsNothing(RoleDefinitions) Then
                    Dim tmpTeamID As Integer = -1
                    Dim tmpRoleID As Integer = RoleDefinitions.parseRoleNameID(_q2, tmpTeamID)
                    If tmpRoleID > 0 Then
                        tmpResult = RoleDefinitions.getRoleDefByID(tmpRoleID).name
                    End If
                End If
            End If
            q2Bezeichner = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' besetzt die aktuelle PTTCompInfo mit den Daten aus dem Shape
    ''' </summary>
    ''' <param name="pptShape"></param>
    Public Sub getValuesFromPPTShape(ByVal pptShape As Microsoft.Office.Interop.PowerPoint.Shape)

        Try

            If Not IsNothing(pptShape) Then


                With pptShape

                    If .Tags.Item("BID").Length > 0 Then
                        bigType = CType(.Tags.Item("BID"), ptReportBigTypes)
                    End If

                    If .Tags.Item("DID").Length > 0 Then
                        detailID = CType(.Tags.Item("DID"), ptReportComponents)
                    End If

                    If .Tags.Item("PNM").Length > 0 Then
                        _pName = .Tags.Item("PNM")
                    End If

                    If .Tags.Item("VNM").Length > 0 Then
                        _vName = .Tags.Item("VNM")
                    Else
                        _vName = ""
                    End If

                    If .Tags.Item("PRPF").Length > 0 Then
                        _prPF = CType(.Tags.Item("PRPF"), ptPRPFType)
                    End If

                    If .Tags.Item("VPID").Length > 0 Then
                        _vpid = CType(.Tags.Item("VPID"), String)
                    End If

                    If .Tags.Item("Q1").Length > 0 Then
                        elementTyp = CType(.Tags.Item("Q1"), ptElementTypen)
                    End If

                    If .Tags.Item("Q2").Length > 0 Then
                        q2 = .Tags.Item("Q2")
                    End If


                    If .Tags.Item("TXT").Length > 0 Then
                        text = CStr(.Tags.Item("TXT"))
                    End If

                    Dim tmpLD As Date = StartofCalendar
                    If .Tags.Item("SRLD").Length > 0 Then
                        Try
                            tmpLD = CDate(.Tags.Item("SRLD"))
                        Catch ex As Exception

                        End Try

                    End If

                    Dim tmpRD As Date = StartofCalendar
                    If .Tags.Item("SRRD").Length > 0 Then
                        Try
                            tmpRD = CDate(.Tags.Item("SRRD"))
                        Catch ex As Exception

                        End Try

                    End If

                    If ((getColumnOfDate(tmpRD) > getColumnOfDate(tmpLD)) And (tmpLD > StartofCalendar)) Then
                        zeitRaumLeft = tmpLD
                        zeitRaumRight = tmpRD
                    End If



                End With

            End If


        Catch ex As Exception
            Dim a As Integer = 1
        End Try
    End Sub
End Class
