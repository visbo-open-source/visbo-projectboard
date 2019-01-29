Public Class clsSmartPPTChartInfo


    Public Property hproj As clsProjekt
    Public Property hproj2 As clsProjekt
    Public Property vglProj As clsProjekt
    Public Property bigType As ptReportBigTypes

    Public Property detailID As PTprdk

    Public Property chartTyp As PTChartTypen
    Public Property vergleichsArt As PTVergleichsArt
    Public Property vergleichsTyp As PTVergleichsTyp
    Public Property vergleichsDatum As Date
    Public Property einheit As PTEinheiten
    Public Property elementTyp As ptElementTypen
    Public Property q2 As String

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
                pName = ""
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
                vName = ""
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
                prPF = ptPRPFType.project
            End If
        End Get
        Set(value As ptPRPFType)
            If IsNothing(_hproj) Then
                _prPF = ptPRPFType.project
            End If
        End Set
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
    ''' besetzt die aktuelle PTTChartInfo mit den Daten aus der Shape
    ''' </summary>
    ''' <param name="pptShape"></param>
    Public Sub getValuesFromPPTShape(ByVal pptShape As Microsoft.Office.Interop.PowerPoint.Shape)
        Try

            If Not IsNothing(pptShape) Then

                If pptShape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    Dim pptChart As Microsoft.Office.Interop.PowerPoint.Chart = pptShape.Chart

                    With pptShape

                        If .Tags.Item("CHT").Length > 0 Then
                            chartTyp = CType(.Tags.Item("CHT"), PTChartTypen)
                        End If

                        If .Tags.Item("ASW").Length > 0 Then
                            einheit = CType(.Tags.Item("ASW"), PTEinheiten)
                        End If

                        If .Tags.Item("VGLA").Length > 0 Then
                            vergleichsArt = CType(.Tags.Item("VGLA"), PTVergleichsArt)
                        End If

                        If .Tags.Item("VGLT").Length > 0 Then
                            vergleichsTyp = CType(.Tags.Item("VGLT"), PTVergleichsTyp)
                        End If


                        If .Tags.Item("VGLD").Length > 0 Then
                            vergleichsDatum = CDate(.Tags.Item("VGLD"))
                        End If


                        If .Tags.Item("Q1").Length > 0 Then
                            elementTyp = CType(.Tags.Item("Q1"), ptElementTypen)
                        End If


                        If .Tags.Item("Q2").Length > 0 Then
                            q2 = .Tags.Item("Q2")
                        End If


                        If .Tags.Item("BID").Length > 0 Then
                            bigType = CType(.Tags.Item("BID"), ptReportBigTypes)
                        End If

                        If .Tags.Item("DID").Length > 0 Then
                            detailID = CType(.Tags.Item("DID"), PTprdk)
                        End If

                        If .Tags.Item("PNM").Length > 0 Then
                            _pName = .Tags.Item("PNM")
                        End If

                        If .Tags.Item("VNM").Length > 0 Then
                            _vName = .Tags.Item("VNM")
                        End If

                        If .Tags.Item("PRPF").Length > 0 Then
                            _prPF = CType(.Tags.Item("PRPF"), ptPRPFType)
                        End If

                    End With

                End If

            End If

        Catch ex As Exception
            Dim a As Integer = 1
        End Try


    End Sub


    Public Sub New()
        _pName = ""
        _vName = ""
        _prPF = ptPRPFType.project
        _hproj = Nothing
        _hproj2 = Nothing
        _vglProj = Nothing
        _bigType = ptReportBigTypes.charts
        _detailID = PTprdk.KostenBalken
        _chartTyp = PTChartTypen.Balken
        _vergleichsArt = PTVergleichsArt.beauftragung
        _vergleichsTyp = PTVergleichsTyp.letzter
        _vergleichsDatum = Date.MinValue
        _einheit = PTEinheiten.euro
        _elementTyp = ptElementTypen.roles
        _q2 = ""
    End Sub
End Class
