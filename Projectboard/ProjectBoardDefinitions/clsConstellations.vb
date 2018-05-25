Public Class clsConstellations
    Private _allConstellations As SortedList(Of String, clsConstellation)

    Public ReadOnly Property Count As Integer

        Get
            Count = _allConstellations.Count
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
    ''' </summary>
    ''' <param name="item"></param>
    ''' <remarks></remarks>
    Public Sub update(item As clsConstellation)

        If Me._allConstellations.ContainsKey(item.constellationName) Then
            Me._allConstellations.Remove(item.constellationName)
        End If

        Me._allConstellations.Add(item.constellationName, item)
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

    End Sub

End Class
