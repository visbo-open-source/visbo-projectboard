Public Class clsAddElements

    Private _name As String
    Private _elementListe As List(Of clsAddElementRules)

    Public Property name As String
        Get
            name = _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    ''' <summary>
    ''' liefert true zurück, wenn die Liste von zu erzeugenden Elementen bereits das Element mit 
    ''' Namen und entsprechenden IsPhase Kennzeichnung enthält
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsElement(ByVal name As String, ByVal isPhase As Boolean) As Boolean
        Get
            Dim found As Boolean = False
            Dim i As Integer = 1
            Do While i <= _elementListe.Count And Not found
                If _elementListe.ElementAt(i - 1).name = name And _
                    _elementListe.ElementAt(i - 1).elemToCreateIsPhase = isPhase Then
                    found = True
                Else
                    i = i + 1
                End If
            Loop

            containsElement = found

        End Get
    End Property

    ''' <summary>
    ''' gibt das Element zurück, das mit Name und Kennzeichen isPhase bezeichnet ist
    ''' Nothing, wenn es nicht existiert
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="isPhase"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getItem(ByVal name As String, ByVal isPhase As Boolean) As clsAddElementRules
        Get
            Dim found As Boolean = False
            Dim i As Integer = 1
            Do While i <= _elementListe.Count And Not found
                If _elementListe.ElementAt(i - 1).name = name And _
                    _elementListe.ElementAt(i - 1).elemToCreateIsPhase = isPhase Then
                    found = True
                Else
                    i = i + 1
                End If
            Loop

            If found Then
                getItem = _elementListe.ElementAt(i - 1)
            Else
                getItem = Nothing
            End If

        End Get
    End Property

    Public Sub addElem(ByVal newElem As clsAddElementRules, ByVal isPhase As Boolean)
        If Me.containsElement(newElem.name, isPhase) Then
            Throw New ArgumentException("Element für Regeln existiert schon ..")
        Else
            _elementListe.Add(newElem)
        End If
    End Sub
    ''' <summary>
    ''' fügt der Regel-Liste eine neue Regel hinzu
    ''' Es muss dazu bereits mind. eine Rule für das Element geben 
    ''' </summary>
    ''' <param name="newRule"></param>
    ''' <remarks></remarks>
    Public Sub addRule(ByVal newRule As clsAddElementRuleItem, ByVal isPhase As Boolean)
        Dim index As Integer = 1
        Dim elemName As String = newRule.newElemName
        Dim elem As clsAddElementRules

        If Not Me.containsElement(elemName, isPhase) Then
            Throw New ArgumentException("Element existiert noch nicht")
        Else
            elem = Me.getItem(elemName, isPhase)
            ' hier wird jetzt das neue Rule-Item in die RuleListe eingefügt 
            elem.add(newRule)
        End If


    End Sub

    Public ReadOnly Property getRule(ByVal index As Integer) As clsAddElementRules
        Get
            If index >= 1 And index <= _elementListe.Count Then
                getRule = _elementListe.ElementAt(index - 1)
            Else
                getRule = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt die x.te Regel für das Element name zurück
    ''' gibt Nothing zurück, wenn es keine x.te Regel für Element name gibt 
    ''' </summary>
    ''' <param name="elemName"></param>
    ''' <param name="lfdNr"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRule(ByVal elemName As String, ByVal isPhase As Boolean, _
                                     ByVal lfdNr As Integer) As clsAddElementRuleItem
        Get
            Dim elem As clsAddElementRules = Me.getItem(elemName, isPhase)

            If IsNothing(elem) Then
                getRule = Nothing
            Else
                If lfdNr <= elem.count And lfdNr >= 1 Then
                    getRule = elem.getItem(lfdNr)
                Else
                    getRule = Nothing
                End If
            End If


        End Get
    End Property


    ''' <summary>
    ''' gibt die Anzahl Regeln zurück, die es für das Element name gibt 
    ''' 0 , wenn es keine einzige Regel gibt 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAnzahlRulesForElem(ByVal name As String, ByVal isPhase As Boolean) As Integer
        Get
            Dim elem As clsAddElementRules = Me.getItem(name, isPhase)

            If IsNothing(elem) Then
                getAnzahlRulesForElem = elem.count
            Else
                getAnzahlRulesForElem = 0
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl Elemente in der Regelliste zuürck
    ''' kann Mehrfach-Nennungen enthalten , deshalb ist die Anzahl der newElem oft kleiner 
    ''' als die Gesamtzahl  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property count() As Integer
        Get
            count = _elementListe.Count
        End Get
    End Property

    Public Sub New()
        _elementListe = New List(Of clsAddElementRules)
        _name = ""
    End Sub
End Class
