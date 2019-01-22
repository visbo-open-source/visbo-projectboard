Public Class clsConstellations
    Private _allConstellations As SortedList(Of String, clsConstellation)
    ' der bool'sche Wert hat aktuell keine Bedeutung; später evtl benutzen um zu bestimmen, das Portfolio Budget aus den Einzelbudgets zu berechnen ist
    Private _listOfLoadedSessionPortfolios As SortedList(Of String, Boolean)

    Public ReadOnly Property Count As Integer

        Get
            Count = _allConstellations.Count
        End Get

    End Property

    ''' <summary>
    ''' fügt der Liste an loadedSessionPortfolios ein Portfolio hinzu 
    ''' wenn zum Portfolio noch kein Summary Projekt existiert, wird es erstellt 
    ''' </summary>
    ''' <param name="portfolioName"></param>
    ''' <returns></returns>
    Public Function addToLoadedSessionPortfolios(ByVal portfolioName As String) As Boolean
        Dim tmpResult As Boolean = True

        If Me.Contains(portfolioName) Then
            If Not _listOfLoadedSessionPortfolios.ContainsKey(portfolioName) Then
                _listOfLoadedSessionPortfolios.Add(portfolioName, True)
            End If
        Else
            tmpResult = False
        End If
        ' was ist mit dem entsprechenden Summary Projekt ... 
        Dim tmpVariantName As String = ""
        If awinSettings.loadPFV Then
            tmpVariantName = ptVariantFixNames.pfv.ToString
        End If
        Dim skey As String = calcProjektKey(portfolioName, tmpVariantName)

        Dim hproj As clsProjekt = AlleProjekte.getProject(key:=skey)
        If Not IsNothing(hproj) Then
            If AlleProjektSummaries.Containskey(skey) Then
                AlleProjektSummaries.Remove(skey, False)
            End If
            AlleProjektSummaries.Add(hproj, updateCurrentConstellation:=False, checkOnConflicts:=False)
        End If

        addToLoadedSessionPortfolios = tmpResult
    End Function

    ''' <summary>
    ''' setzt die Liste der geladenen Session Portfolios zurück 
    ''' </summary>
    Public Sub clearLoadedPortfolios()
        _listOfLoadedSessionPortfolios.Clear()
        AlleProjektSummaries.Clear(False)
    End Sub

    ''' <summary>
    ''' gibt das Gesamt Budget des Zeitraums im Gesamt-Portfolio zurück 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getBudgetOfLoadedPortfolios() As Double
        Get
            Dim tmpResult As Double = 0.0
            For Each pfKvP As KeyValuePair(Of String, Boolean) In _listOfLoadedSessionPortfolios

                Dim key As String = calcProjektKey(pfKvP.Key, ptVariantFixNames.pfv.ToString)
                Dim hproj As clsProjekt = AlleProjekte.getProject(key)

                ' wenn ein Portfolio nicht über Platzhalter geladen wird, dann werden die Summary Projekte in AlleProjektSummaries platziert ..
                If IsNothing(hproj) Then
                    hproj = AlleProjektSummaries.getProject(key)
                End If

                Dim teilbudget As Double, pk As Double, sk As Double, rk As Double, ergebnis As Double

                If Not IsNothing(hproj) Then
                    Call hproj.calculateRoundedKPI(teilbudget, pk, sk, rk, ergebnis)
                    tmpResult = tmpResult + teilbudget
                End If

            Next
            getBudgetOfLoadedPortfolios = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Szenarien zurück, getrennt durch ;, die pName und vName referenzieren 
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="vname"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSzenarioNamesWith(ByVal pname As String, ByVal vname As String) As String
        Get
            Dim tmpResult As String = ""

            For Each kvp As KeyValuePair(Of String, clsConstellation) In _allConstellations

                If kvp.Key = "Last" Or kvp.Key = "Session" Then
                    ' nichts tun, sind systembedingt 
                Else
                    Dim tmpKey As String = calcProjektKey(pname, vname)
                    If kvp.Value.contains(tmpKey, False) Then
                        If tmpResult = "" Then
                            tmpResult = kvp.Key
                        Else
                            tmpResult = tmpResult & "; " & kvp.Key
                        End If
                    End If
                End If
            Next

            getSzenarioNamesWith = tmpResult

        End Get
    End Property
    Public ReadOnly Property Liste As SortedList(Of String, clsConstellation)

        Get
            Liste = _allConstellations
        End Get

    End Property

    Public ReadOnly Property getConstellation(name As String) As clsConstellation
        Get

            If _allConstellations.ContainsKey(name) Then
                getConstellation = _allConstellations.Item(name)
            Else
                getConstellation = Nothing
            End If

        End Get
    End Property

    Public ReadOnly Property Contains(name As String) As Boolean
        Get
            Contains = _allConstellations.ContainsKey(name)
        End Get
    End Property

    ''' <summary>
    ''' trägt die angegebene Constellation in die Liste der ProjektConstellations ein. 
    ''' </summary>
    ''' <param name="item">Constellation</param>
    Sub Add(ByVal item As clsConstellation)

        Try
            _allConstellations.Add(item.constellationName, item)
        Catch ex As Exception
            Dim errmsg = "Program-/Portfolio Name existiert bereits: " & item.constellationName
            Throw New ArgumentException(errmsg)
        End Try


    End Sub


    ''' <summary>
    ''' aktualisiert in jeder Constellation die Variante mit Namen oldvName mit dem Namen newvName
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="oldvName"></param>
    ''' <param name="newvName"></param>
    ''' <remarks></remarks>
    Public Sub updateVariantName(ByVal pName As String, ByVal oldvName As String, ByVal newvName As String)

        For Each kvp As KeyValuePair(Of String, clsConstellation) In _allConstellations
            Call kvp.Value.updateVariantName(pName, oldvName, newvName)
        Next

    End Sub


    ''' <summary>
    ''' ersetzt oder fügt eine neue Konstellation mit dem Namen ein 
    ''' das wird die neue Konstellation , das heisst es muss auch ein neues Summary Projekt für diese Konstellation gemacht werden  
    ''' </summary>
    ''' <param name="item"></param>
    ''' <remarks></remarks>
    Public Sub update(item As clsConstellation)

        Me.clearLoadedPortfolios()

        If Me._allConstellations.ContainsKey(item.constellationName) Then
            Me._allConstellations.Remove(item.constellationName)
        End If

        Me._allConstellations.Add(item.constellationName, item)

        ' jetzt das in loadedSessionPortfolios reinbringen 
        Me.addToLoadedSessionPortfolios(item.constellationName)

    End Sub

    Sub Remove(ByVal key As String)

        Try
            _allConstellations.Remove(key)
        Catch ex As Exception
            Throw New ArgumentException("Konstellation" & " key " & "konnte nicht gelöscht werden ")
        End Try

    End Sub

    Sub New()

        _allConstellations = New SortedList(Of String, clsConstellation)
        _listOfLoadedSessionPortfolios = New SortedList(Of String, Boolean)

    End Sub

End Class
