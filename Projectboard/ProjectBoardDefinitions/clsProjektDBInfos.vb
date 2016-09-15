
Imports System.Collections.Generic


''' <summary>
''' Klasse enthält die für das Administrationstool notwendigen Informationen zu den einzelnen Projekten
''' Projektname, Varianten-Name, Timestamp, verantwortlicher
''' Schlüssel ist Projektname#VariantenName#Verantwortlicher
''' Value ist eine Sortierte Liste von Timestamps für diesen Schlüssel 
''' </summary>
''' <remarks></remarks>
Public Class clsProjektDBInfos
    '  
    ' 
    ' Sortierte 

    'Private _pliste As SortedList(Of String, clsProjektDBInfo)

    Private _pliste As SortedList(Of String, SortedList(Of Date, String))

    ''' <summary>
    ''' löscht die ProjektListe
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clear()

        _pliste.Clear()

    End Sub

    Public Property Liste() As SortedList(Of String, SortedList(Of Date, String))
        Get
            Liste = _pliste
        End Get
        Set(value As SortedList(Of String, SortedList(Of Date, String)))
            _pliste = value
        End Set
    End Property
    ''' <summary>
    ''' gibt die Anzahl der Liste-Elemente (Projekte)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count() As Integer
        Get
            Count = _pliste.Count
        End Get
    End Property
    ''' <summary>
    ''' gibt die Anzahl der Listen-Elemente (TimeStamps)
    ''' </summary>
    ''' <param name="projektKey">Projektname#variantName</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count(ByVal projektKey As String) As Integer
        Get
            If _pliste.ContainsKey(projektKey) Then

                Count = _pliste.Item(projektKey).Count
            Else
                Count = 0
            End If
        End Get
    End Property
    ''' <summary>
    ''' gibt für einen gegebenen Projekt-Namen eine sortierte Liste of Timestamps, Projektschlüssel zurück  
    ''' </summary>
    ''' <param name="projektKey">Projektname#variantName</param>
    ''' <value></value>
    ''' <returns>sortierte Liste of Dates </returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTimeStamps(ByVal projektKey As String) As SortedList(Of Date, String)
        Get

            Dim ergebnis As New SortedList(Of Date, String)

            If _pliste.ContainsKey(projektKey) Then
                ergebnis = _pliste.Item(projektKey)
            End If
            getTimeStamps = ergebnis

        End Get

    End Property
    ''' <summary>
    ''' fügt eine komplette Projekthistorie der ProjektDBInfo hinzu 
    ''' </summary>
    ''' <param name="pHistorie">eine nicht-leere Projekt-Historie</param>
    ''' <remarks>Exception, wenn leere Projekthistorie übergeben wird</remarks>
    Public Sub Add(ByVal pHistorie As clsProjektHistorie)

        Dim hproj As clsProjekt
        Dim suchString As String
        Dim tmpHistListe As SortedList(Of Date, String)
        Dim found As Boolean


        If pHistorie.Count = 0 Then
            Throw New ArgumentException("keine Historie vorhanden !")
        End If

        hproj = pHistorie.First

        suchString = calcProjektKey(hproj)

        If _pliste.ContainsKey(suchString) Then
            tmpHistListe = _pliste.Item(suchString)
            found = True
        Else
            tmpHistListe = New SortedList(Of Date, String)
            found = False
        End If

        For Each kvp As KeyValuePair(Of Date, clsProjekt) In pHistorie.liste

            suchString = calcProjektKey(kvp.Value)
            tmpHistListe.Add(kvp.Key, suchString)
        Next

        ' wenn die Liste bereits existiert hat 
        If found Then
            _pliste.Remove(suchString)
        End If

        _pliste.Add(suchString, tmpHistListe)

    End Sub

    ''' <summary>
    ''' fügt der Liste von Projekten mit Timestamps das Paar Projektkey, Timestamp zu
    ''' wenn der Projektkey schon existiert: einsortieren des Timestamps in die sortierte liste von timestamps des Projektes 
    ''' wenn der Projektkey noch nicht existiert: anlegen und anlegen der vorerst nur aus einem Eintrag bestehenden Timestamp Liste 
    ''' </summary>
    ''' <param name="projektKey">setzt sich zusammen aus pName#variantName</param>
    ''' <param name="timestamp"></param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal projektKey As String, ByVal timestamp As Date)

        Dim tmpHistListe As SortedList(Of Date, String)


        If _pliste.ContainsKey(projektKey) Then

            If _pliste.Item(projektKey).ContainsKey(timestamp) Then
                ' nichts tun 
            Else
                _pliste.Item(projektKey).Add(timestamp, projektKey)
            End If

        Else
            ' existiert noch nicht - dann neu aufnehmen 
            tmpHistListe = New SortedList(Of Date, String)
            tmpHistListe.Add(timestamp, projektKey)
            _pliste.Add(projektKey, tmpHistListe)
        End If


    End Sub

    ''' <summary>
    ''' löscht aus der Projekt-/Timestamp Struktur das Projekt mit 
    ''' Schlüssel pName#variantName  
    ''' </summary>
    ''' <param name="projektkey"></param>
    ''' <remarks></remarks>
    Public Sub Remove(ByVal projektkey As String)

        If _pliste.ContainsKey(projektkey) Then

            _pliste.Remove(projektkey)

        Else
            ' nichts tun - existiert sowieso nicht 
        End If


    End Sub

    ''' <summary>
    ''' löscht in der Projekt-/Timestamp Struktur den Eintrag projektkey/timeStamp 
    ''' </summary>
    ''' <param name="projektkey">pName#variantName</param>
    ''' <param name="timestamp">Zeitstempel, wann das Projekt gespeichert wurde</param>
    ''' <remarks></remarks>
    Public Sub Remove(ByVal projektkey As String, ByVal timestamp As Date)

        If _pliste.ContainsKey(projektkey) Then

            If _pliste.Item(projektkey).ContainsKey(timestamp) Then
                _pliste.Item(projektkey).Remove(timestamp)
            End If

            If _pliste.Item(projektkey).Count = 0 Then
                ' es gibt keine Timestamps mehr - also kann der Eintrag für den Projektschlüssel ganz gelöscht werden 
                _pliste.Remove(projektkey)
            End If

        Else
            ' nichts tun - existiert sowieso nicht 

        End If


    End Sub

    Public Sub New()
        _pliste = New SortedList(Of String, SortedList(Of Date, String))
    End Sub

    ' ''' <summary>
    ' ''' Klasse, die in der Klasse clsProjektDBInfos verwendet wird 
    ' ''' Sortierte Liste mit Datum als Suchschlüssel; value ist Platzhalter boolean; die werte wären in diesem Falle alle gleich 
    ' ''' nämlich pName#variantName#verantwortlicher  
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Private Class clsProjektDBInfo


    '    Private _histListe As SortedList(Of Date, String)

    '    Public ReadOnly Property historyListe() As SortedList(Of Date, String)
    '        Get

    '            historyListe = _histListe

    '        End Get
    '    End Property

    '    Public Sub add(ByVal timeStamp As Date, ByVal projektKey As String)

    '        If _histListe.ContainsKey(timeStamp) Then
    '            ' do nothing 
    '        Else
    '            _histListe.Add(timeStamp, projektKey)
    '        End If

    '    End Sub


    '    Public Sub New()
    '        _histListe = New SortedList(Of Date, String)
    '    End Sub


    'End Class
End Class
