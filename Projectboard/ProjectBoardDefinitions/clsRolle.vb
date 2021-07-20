Imports System.Math
Public Class clsRolle

    ' die ID der Rolle
    Private _uid As Integer

    ' der Monats-Array mit den monatlichen Bedarfen
    Private _bedarf() As Double

    ' tk 24.11.18 isCalculated wird nie gebraucht 
    'Private _isCalculated As Boolean

    ' neu dazu gekommen: die evtl der Rolle zugeordnete TeamID ; -1 heisst es gibt keine 
    Private _teamID As Integer

    Public Property teamID As Integer
        Get
            teamID = _teamID
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then

                If value = -1 Then
                    _teamID = value

                ElseIf RoleDefinitions.containsUid(value) Then
                    _teamID = value

                End If

            End If
        End Set
    End Property

    ''' <summary>
    ''' gibt die NAmeID zurück, in der Form uid.toString; teamID.to string bzw uid.tostring
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getNameID As String
        Get
            getNameID = RoleDefinitions.bestimmeRoleNameID(_uid, _teamID)
        End Get
    End Property

    ''' <summary>
    ''' bestimmt, ob die Rolle identisch mit der übergebenen Rolle ist 
    ''' </summary>
    ''' <param name="vRolle"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vRolle As clsRolle) As Boolean
        Get
            Dim stillOK As Boolean = False



            If Me.uid = vRolle.uid And
                    Not arraysAreDifferent(Me.Xwerte, vRolle.Xwerte) And
                   Me.teamID = vRolle.teamID Then
                stillOK = True
            Else
                stillOK = False
            End If

            isIdenticalTo = stillOK
        End Get
    End Property
    'Public Property isCalculated() As Boolean
    '    Get
    '        isCalculated = _isCalculated
    '    End Get
    '    Set(value As Boolean)
    '        _isCalculated = value
    '    End Set
    'End Property

    Public Property uid() As Integer
        Get

            uid = _uid

        End Get

        Set(value As Integer)

            If Not IsNothing(value) Then
                If RoleDefinitions.containsUid(value) Then
                    _uid = value
                End If
            End If

        End Set
    End Property
    '
    '
    '
    Public ReadOnly Property getDimension As Integer
        Get
            getDimension = _bedarf.Length - 1
        End Get
    End Property

    Public Property Xwerte() As Double()
        Get
            Xwerte = _bedarf
        End Get

        Set(values As Double())

            Dim ub As Integer = UBound(values)
            Dim tmpArray() As Double
            ReDim tmpArray(ub)

            For i As Integer = 0 To ub
                tmpArray(i) = values(i)
            Next
            _bedarf = tmpArray

        End Set

    End Property

    ''' <summary>
    ''' gibt die Summe zurück bis zum angegebenen Index einschließlich
    ''' kann verwendet werden, um die actualdata.sum für die Rolle zu bestimmen ...
    ''' </summary>
    ''' <param name="index"></param>
    ''' <returns></returns>
    Public ReadOnly Property getSumUntil(ByVal index As Integer) As Double
        Get

            Dim ergebnis As Double = 0.0
            If index < 0 Or index > _bedarf.Length - 1 Then
                ergebnis = 0.0
            Else
                For i As Integer = 0 To index
                    ergebnis = ergebnis + _bedarf(i)
                Next
            End If

            getSumUntil = ergebnis
        End Get
    End Property

    ''' <summary>
    ''' gibt die Summe zurück ab index+1 bis Ende des Arrays
    ''' kann verwendet werden, um die forecast.sum für die Rolle zu bestimmen ...
    ''' </summary>
    ''' <param name="index"></param>
    ''' <returns></returns>
    Public ReadOnly Property getSumFrom(ByVal index As Integer) As Double
        Get
            Dim ergebnis As Double = 0.0
            If index < 0 Then
                ' damit das dann nachher einfach alles aufsummiert ...
                index = -1
            End If

            If index > _bedarf.Length - 1 Then
                ergebnis = 0.0
            Else
                For i As Integer = index + 1 To _bedarf.Length - 1
                    ergebnis = ergebnis + _bedarf(i)
                Next
            End If

            getSumFrom = ergebnis
        End Get
    End Property

    Public Property Xwerte(ByVal index As Integer) As Double

        Get
            Xwerte = _bedarf(index)
        End Get

        Set(value As Double)
            _bedarf(index) = value
        End Set

    End Property
    '
    '
    '
    Public ReadOnly Property name() As String

        Get

            name = RoleDefinitions.getRoleDefByID(_uid).name

        End Get

    End Property

    Public ReadOnly Property teamName() As String
        Get
            Dim tmpResult As String = ""
            If _teamID > 0 Then
                If RoleDefinitions.containsUid(_teamID) Then
                    tmpResult = RoleDefinitions.getRoleDefByID(_teamID).name
                End If
            End If

            teamName = tmpResult
        End Get
    End Property
    '
    '
    '
    Public ReadOnly Property farbe() As Object

        Get

            farbe = visboFarbeBlau
            ' tk 20.7.21 keine individuelle Farbe mehr für Rollen 
            'farbe = RoleDefinitions.getRoleDefByID(_uid).farbe

        End Get

    End Property
    '
    '
    '
    Public ReadOnly Property Startkapa() As Double

        Get

            Startkapa = RoleDefinitions.getRoleDefByID(_uid).defaultKapa

        End Get


    End Property

    ''' <summary>
    ''' gibt zurück, ob es sich bei der Rolle um eine externe Ressource handelt ...
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property isExtern() As Boolean
        Get
            isExtern = RoleDefinitions.getRoleDefByID(_uid).isExternRole
        End Get
    End Property

    ''' <summary>
    ''' gets day-rate according to skill / role
    ''' Use skill rate when skill is given and role is combinedRole
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property tagessatzIntern() As Double

        Get
            ' wenn eine Skill angegeben ist bei einer Orga-Unit Sammelrolle: nimm Tagessatz der Skill 
            ' wenn eine Skill angegeben ist bei einer Person (also ohne Kinder): nimm Tagessatz der Person
            ' wenn keine Skill angegeben ist: nimm Tagessatz der Orga-Unit / Skill 

            Try
                If _teamID > 0 Then
                    If RoleDefinitions.getRoleDefByID(_uid).isCombinedRole Then
                        ' use day-rate of skill
                        tagessatzIntern = RoleDefinitions.getRoleDefByID(_teamID).tagessatzIntern
                    Else
                        ' use day-rate of orga-unit / skill 
                        tagessatzIntern = RoleDefinitions.getRoleDefByID(_uid).tagessatzIntern
                    End If
                Else
                    tagessatzIntern = RoleDefinitions.getRoleDefByID(_uid).tagessatzIntern
                End If
            Catch ex As Exception
                tagessatzIntern = RoleDefinitions.getRoleDefByID(_uid).tagessatzIntern
            End Try



        End Get

    End Property

    '
    '
    '
    Public ReadOnly Property summe() As Double

        Get

            summe = _bedarf.Sum


        End Get

    End Property
    '

    Public Sub CopyTo(ByRef newrole As clsRolle)

        With newrole
            .uid = _uid
            .teamID = _teamID
            .Xwerte = _bedarf
        End With

    End Sub

    ''' <summary>
    ''' Sub berechnet die neuen Werte so, daß die Charakterisitik der Werte möglichst erhalten bleibt 
    ''' Übergeben wird die neue Länge - es wird dann entschieden, welche Charakteristik am ehesten zutrifft - danach werden die Werte neu bestimmt
    ''' newlength ist die echte länge, also z.Bsp steht 2 für 2 Monate 
    ''' </summary>
    ''' <param name="newLength"></param>
    ''' <remarks></remarks>
    Public Sub adjustLength(ByVal newLength As Integer)
        Dim oldLength As Integer
        Dim oldSum As Double, newSum As Double
        Dim newValues() As Double
        Dim diff As Integer
        Dim typus As Integer

        Dim ix As Integer, i As Integer, lefti As Integer, righti As Integer
        Dim anzRechts As Integer, anzLinks As Integer
        Dim rValues() As Double, lValues() As Double
        Dim notCorrect As Boolean


        ReDim newValues(newLength - 1)
        oldLength = UBound(_bedarf) + 1

        ' wenn keine Änderung vorzunehmen ist, dann Exit ... 
        If newLength = oldLength Then
            Exit Sub
        End If

        oldSum = 0.0
        For i = 0 To oldLength - 1
            oldSum = oldSum + _bedarf(i)
        Next

        Dim avg As Double
        Dim min As Double, max As Double

        avg = Round(oldSum / oldLength, 0)
        min = _bedarf.Min
        max = _bedarf.Max

        newSum = newLength / oldLength * oldSum
        typus = definecharacteristics(min, max, avg)

        If oldLength < newLength Then
            ' verlängern ... 

            diff = newLength - oldLength
            Dim korrfaktor As Double
            korrfaktor = Round(diff * oldSum / oldLength - diff * avg, 0)

            i = 0
            Dim found As Boolean = False

            Select Case typus
                Case 1
                    ' aufsteigend ...
                    While i <= oldLength - 2 And Not found
                        If _bedarf(i) < avg And _bedarf(i + 1) >= avg Then
                            found = True
                        Else
                            i = i + 1
                        End If
                    End While
                    ' jetzt werden die neuen Werte eingefügt 
                    For ix = 0 To i
                        newValues(ix) = _bedarf(ix)
                    Next ix
                    For ix = i + 1 To i + diff
                        newValues(ix) = avg
                    Next ix
                    If korrfaktor > 0 Then
                        newValues(i + diff) = newValues(i + diff) + korrfaktor
                    End If

                    For ix = i + diff + 1 To newLength
                        newValues(ix) = _bedarf(ix - diff)
                    Next ix
                Case 2
                    ' die Buckel Funktion 
                    If min = max Then
                        ' es ist einfach - nur Felder ergänzen ...

                        For ix = 0 To oldLength
                            newValues(ix) = _bedarf(ix)
                        Next ix

                        For ix = oldLength + 1 To newLength
                            newValues(ix) = avg
                        Next ix
                    Else
                        ' jetzt muss im linken Teil und im rechten Teil abwechselnd ergänzt werden 
                        lefti = 0
                        While lefti <= oldLength - 2 And Not found
                            If _bedarf(lefti) < avg And _bedarf(lefti + 1) >= avg Then
                                found = True
                            Else
                                lefti = lefti + 1
                            End If
                        End While

                        righti = oldLength
                        While righti >= 1 And Not found
                            If _bedarf(righti) < avg And _bedarf(righti - 1) >= avg Then
                                found = True
                            Else
                                righti = righti - 1
                            End If
                        End While
                        If lefti > righti Then
                            Call MsgBox("Fehler in clsRolle, adjustLength, verlängern, case 3")
                            lefti = righti
                        End If
                        ' bestimme Werte, die links bzw rechts ergänzt werden müssen ....
                        i = diff
                        Dim lefthand As Boolean = True
                        anzLinks = 0
                        anzRechts = 0
                        ReDim lValues(diff)
                        ReDim rValues(diff)
                        While i > 0
                            If lefthand Then
                                lefthand = False
                                anzLinks = anzLinks + 1
                                lValues(anzLinks) = avg

                            Else
                                lefthand = True
                                anzRechts = anzRechts + 1
                                rValues(anzRechts) = avg

                            End If
                            i = i - 1
                        End While

                        lValues(anzLinks) = lValues(anzLinks) + korrfaktor

                        For ix = 0 To lefti
                            newValues(ix) = _bedarf(ix)
                        Next
                        For ix = lefti + 1 To lefti + anzLinks
                            newValues(ix) = lValues(ix - lefti)
                        Next
                        For ix = lefti + 1 To righti
                            newValues(ix + anzLinks) = _bedarf(ix)
                        Next
                        For ix = righti + anzLinks + 1 To righti + anzLinks + anzRechts
                            newValues(ix) = rValues(ix - righti - anzLinks)
                        Next

                        For ix = righti + anzLinks + anzRechts + 1 To anzLinks + anzRechts + oldLength - 1
                            newValues(ix) = _bedarf(ix - anzLinks - anzRechts)
                        Next

                    End If
                Case 3
                    ' absteigend ...
                    While i <= oldLength - 2 And Not found
                        If _bedarf(i) > avg And _bedarf(i + 1) <= avg Then
                            found = True
                        Else
                            i = i + 1
                        End If
                    End While
                    ' jetzt werden die neuen Werte eingefügt 
                    For ix = 0 To i
                        newValues(ix) = _bedarf(ix)
                    Next ix

                    For ix = i + 1 To i + diff
                        newValues(ix) = avg
                    Next ix

                    If korrfaktor > 0 Then
                        newValues(i + diff) = newValues(i + diff) + korrfaktor
                    End If

                    For ix = i + diff + 1 To newLength
                        newValues(ix) = _bedarf(ix - diff)
                    Next ix

            End Select

        ElseIf oldLength > newLength Then
            ' verkürzen
            ' es werden in der Mitte Anzahl <Diff> Werte herausgenommen ;
            ' über den Korektur Faktor wird ausgeglichen, daß die ZielSumme wieder annähernd stimmt 
            '

            diff = oldLength - newLength
            Dim korrfaktor As Double
            Dim abzug As Double = _bedarf(i)
            Dim righthand As Boolean = True
            Dim tmpWert As Integer = diff

            i = 0
            Dim found As Boolean = False

            Select Case typus
                Case 1
                    ' aufsteigend ...
                    While i <= oldLength - 2 And Not found
                        If _bedarf(i) < avg And _bedarf(i + 1) >= avg Then
                            found = True
                        Else
                            i = i + 1
                        End If
                    End While
                    ' jetzt werden die neuen Werte aufgebaut 
                    lefti = i - 1
                    righti = i + 1



                    While tmpWert > 1
                        If righthand Then
                            righthand = False
                            If righti + 1 <= oldLength Then
                                righti = righti + 1
                                abzug = abzug + _bedarf(righti - 1)
                            Else
                                lefti = lefti - 1
                                If lefti < -1 Then
                                    Call MsgBox("Fehler in clsRolle, adjustlength, verkürzen 001")
                                    lefti = 0
                                End If
                                abzug = abzug + _bedarf(lefti + 1)
                            End If
                        Else
                            righthand = True
                            If lefti >= 0 Then
                                lefti = lefti - 1
                                abzug = abzug + _bedarf(lefti + 1)
                                righthand = True
                            Else
                                righti = righti + 1
                                If righti > oldLength Then
                                    Call MsgBox("Fehler in clsRolle, adjustlength, verkürzen 002")
                                    righti = oldLength
                                End If
                                abzug = abzug + _bedarf(righti - 1)
                            End If

                        End If
                        tmpWert = tmpWert - 1
                    End While

                    korrfaktor = appInstance.WorksheetFunction.Round(abzug - diff * avg, 0)

                    For ix = 0 To lefti
                        newValues(ix) = _bedarf(ix)
                    Next ix

                    For ix = righti To oldLength - 1
                        newValues(ix - diff) = _bedarf(ix)
                    Next ix



                    If korrfaktor > 0 Then
                        ix = lefti
                        newValues(ix) = newValues(ix) + korrfaktor
                        ' jetzt werden evtl vorhandene Peaks nach rechts geglättet ...
                        ' so daß weiterhin die Charakteristik "aufsteigend" beibehalten wird

                        While ix < newLength - 1

                            If newValues(ix) > newValues(ix + 1) Then
                                notCorrect = True
                            Else
                                notCorrect = False
                            End If

                            While notCorrect

                                newValues(ix) = newValues(ix) - 1
                                newValues(ix + 1) = newValues(ix + 1) + 1
                                If newValues(ix) > newValues(ix + 1) Then
                                    notCorrect = True
                                Else
                                    notCorrect = False
                                End If


                            End While

                            ix = ix + 1
                        End While

                    Else
                        ' jetzt muss darauf geachtet werden, daß kein Wert durch die Korrektur kleiner als Null werden kann 
                        ix = righti
                        While ix <= newLength - 1 And korrfaktor < 0
                            If newValues(ix) + korrfaktor >= 0.0 Then
                                newValues(ix) = newValues(ix) + korrfaktor
                                korrfaktor = 0
                            Else
                                korrfaktor = korrfaktor + newValues(ix)
                                newValues(ix) = 0
                                ix = ix + 1
                            End If

                        End While


                        ' jetzt werden evtl vorhandene Peaks nach links geglättet ...
                        ' so daß weiterhin die Charakteristik "aufsteigend von links nach rechts" beibehalten wird

                        While ix > 0

                            If newValues(ix) < newValues(ix - 1) Then
                                notCorrect = True
                            Else
                                notCorrect = False
                            End If

                            While notCorrect

                                newValues(ix) = newValues(ix) + 1
                                newValues(ix - 1) = newValues(ix - 1) - 1
                                If newValues(ix) < newValues(ix - 1) Then
                                    notCorrect = True
                                Else
                                    notCorrect = False
                                End If


                            End While

                            ix = ix - 1
                        End While

                    End If



                Case 2
                    ' die Buckel Funktion 
                    If min = max Then
                        ' es ist einfach - nur aufbauen ....

                        For ix = 0 To newLength
                            newValues(ix) = _bedarf(ix)
                        Next ix


                    Else
                        ' jetzt muss im linken Teil und im rechten Teil abwechselnd gelöscht werden  
                        lefti = 0
                        While lefti <= oldLength - 2 And Not found
                            If _bedarf(lefti) < avg And _bedarf(lefti + 1) >= avg Then
                                found = True
                            Else
                                lefti = lefti + 1
                            End If
                        End While

                        righti = oldLength
                        While righti >= 1 And Not found
                            If _bedarf(righti) < avg And _bedarf(righti - 1) >= avg Then
                                found = True
                            Else
                                righti = righti - 1
                            End If
                        End While

                        If lefti > righti Then
                            Call MsgBox("Fehler in clsRolle, adjustLength, verkürzen, case 2")
                            lefti = righti
                        End If

                        Dim leftil As Integer = lefti
                        Dim leftir As Integer = lefti
                        Dim rightil As Integer = righti
                        Dim rightir As Integer = righti


                        ' bestimme Werte, die links bzw rechts gelöscht werden müssen ....
                        i = diff
                        Dim lefthand As Boolean = False
                        Dim lefthandRight As Boolean = True
                        Dim righthandLeft As Boolean = True

                        abzug = _bedarf(lefti)
                        leftil = lefti - 1
                        leftir = lefti + 1

                        Dim nothingDone As Boolean = True
                        While i > 1
                            While nothingDone

                                If lefthand Then

                                    If lefthandRight Then

                                        If leftir + 1 <= rightil Then
                                            abzug = abzug + _bedarf(leftir)
                                            leftir = leftir + 1
                                            nothingDone = False
                                        End If

                                        lefthandRight = False
                                    Else
                                        If leftil >= 0 Then
                                            abzug = abzug + _bedarf(leftil)
                                            leftil = leftil - 1
                                            nothingDone = False
                                        End If

                                        lefthandRight = True
                                    End If

                                    lefthand = False

                                Else

                                    If rightil = rightir Then
                                        ' das erste Auftreten ...
                                        abzug = abzug + _bedarf(righti)
                                        rightil = righti - 1
                                        rightir = righti + 1
                                        nothingDone = False


                                    Else
                                        If righthandLeft Then
                                            If rightil - 1 >= leftir Then
                                                abzug = abzug + _bedarf(rightil)
                                                rightil = rightil - 1
                                                nothingDone = False
                                            End If

                                            righthandLeft = False
                                        Else
                                            If rightir <= oldLength - 1 Then
                                                abzug = abzug + _bedarf(rightir)
                                                rightir = rightir + 1
                                                nothingDone = False
                                            Else
                                                If rightil - 1 >= leftir Then
                                                    abzug = abzug + _bedarf(rightil)
                                                    rightil = rightil - 1
                                                    nothingDone = False
                                                End If
                                            End If

                                            righthandLeft = True
                                        End If
                                    End If

                                    lefthand = True

                                End If


                            End While
                            i = i - 1
                            nothingDone = True
                        End While
                        ' jetzt werden die Newvalues aufgebaut ... 
                        Dim nx As Integer = 0

                        For ix = 0 To leftil
                            newValues(nx) = _bedarf(ix)
                            nx = nx + 1
                        Next

                        For ix = leftir To rightil
                            newValues(nx) = _bedarf(ix)
                            nx = nx + 1
                        Next

                        For ix = rightir To oldLength - 1
                            newValues(nx) = _bedarf(ix)
                            nx = nx + 1
                        Next

                        ' jetzt muss die Korrektur vorgenommen werden ...

                        korrfaktor = appInstance.WorksheetFunction.Round(abzug - diff * avg, 0)
                        Dim lx As Integer, rx As Integer
                        lx = CInt(appInstance.WorksheetFunction.Round(newLength - 1 / 4, 0))
                        rx = CInt(appInstance.WorksheetFunction.Round(3 * newLength - 1 / 4, 0))
                        lefthand = True

                        Dim Vorzeichen As Integer
                        If korrfaktor > 0 Then
                            Vorzeichen = 1
                        Else
                            Vorzeichen = -1
                        End If

                        Dim leftvalue As Integer, rightvalue As Integer
                        If korrfaktor * Vorzeichen > 2 Then
                            leftvalue = CInt(appInstance.WorksheetFunction.Round(korrfaktor / 2, 0))
                            rightvalue = leftvalue
                            newValues(lx) = newValues(lx) + leftvalue
                            newValues(rx) = newValues(rx) + rightvalue
                        Else
                            newValues(lx) = newValues(lx) + korrfaktor
                        End If



                    End If
                Case 3

                    ' absteigend ...
                    While i <= oldLength - 2 And Not found
                        If _bedarf(i) >= avg And _bedarf(i + 1) < avg Then
                            found = True
                        Else
                            i = i + 1
                        End If
                    End While
                    ' jetzt werden die neuen Werte aufgebaut 
                    lefti = i - 1
                    righti = i + 1

                    While tmpWert > 1
                        If righthand Then
                            righthand = False
                            If righti + 1 <= oldLength Then
                                righti = righti + 1
                                abzug = abzug + _bedarf(righti - 1)
                            Else
                                lefti = lefti - 1
                                If lefti < -1 Then
                                    Call MsgBox("Fehler in clsRolle, adjustlength, verkürzen 001")
                                    lefti = 0
                                End If
                                abzug = abzug + _bedarf(lefti + 1)
                            End If
                        Else
                            righthand = True
                            If lefti >= 0 Then
                                lefti = lefti - 1
                                abzug = abzug + _bedarf(lefti + 1)
                                righthand = True
                            Else
                                righti = righti + 1
                                If righti > oldLength Then
                                    Call MsgBox("Fehler in clsRolle, adjustlength, verkürzen 002")
                                    righti = oldLength
                                End If
                                abzug = abzug + _bedarf(righti - 1)
                            End If

                        End If
                        tmpWert = tmpWert - 1
                    End While

                    korrfaktor = appInstance.WorksheetFunction.Round(abzug - diff * avg, 0)

                    For ix = 0 To lefti
                        newValues(ix) = _bedarf(ix)
                    Next ix

                    For ix = righti To oldLength - 1
                        newValues(ix - diff) = _bedarf(ix)
                    Next ix



                    If korrfaktor > 0 Then
                        ix = lefti
                        newValues(ix) = newValues(ix) + korrfaktor
                        ' jetzt werden evtl vorhandene Peaks nach links geglättet ...
                        ' so daß weiterhin die Charakteristik "absteigend" beibehalten wird

                        While ix - 1 > 0

                            If newValues(ix - 1) < newValues(ix) Then
                                notCorrect = True
                            Else
                                notCorrect = False
                            End If

                            While notCorrect

                                newValues(ix) = newValues(ix) - 1
                                newValues(ix - 1) = newValues(ix - 1) + 1
                                If newValues(ix - 1) < newValues(ix) Then
                                    notCorrect = True
                                Else
                                    notCorrect = False
                                End If


                            End While

                            ix = ix - 1
                        End While

                    Else
                        ' jetzt muss darauf geachtet werden, daß kein Wert durch die Korrektur kleiner als Null werden kann 
                        ix = righti
                        While ix - 1 > 0 And korrfaktor < 0
                            If newValues(ix) + korrfaktor >= 0.0 Then
                                newValues(ix) = newValues(ix) + korrfaktor
                                korrfaktor = 0
                            Else
                                korrfaktor = korrfaktor + newValues(ix)
                                newValues(ix) = 0
                                ix = ix - 1
                            End If

                        End While


                        ' jetzt werden evtl vorhandene Peaks nach rechts geglättet ...
                        ' so daß weiterhin die Charakteristik "absteigend von links nach rechts" beibehalten wird
                        ix = righti

                        While ix < newLength - 2

                            If newValues(ix) < newValues(ix + 1) Then
                                notCorrect = True
                            Else
                                notCorrect = False
                            End If

                            While notCorrect

                                newValues(ix) = newValues(ix) + 1
                                newValues(ix + 1) = newValues(ix + 1) - 1
                                If newValues(ix) < newValues(ix + 1) Then
                                    notCorrect = True
                                Else
                                    notCorrect = False
                                End If


                            End While

                            ix = ix - 1
                        End While

                    End If

            End Select
        End If

    End Sub

    ''' <summary>
    ''' bestimmt die Charakteristik des Verlaufs: 
    ''' 1-minimum vorne, max hinten -  steigender Verlauf
    ''' 2-Max in der Mitte bzw. einigermaßen konstanter Verlauf
    ''' 3-max vorne, min hinten -  fallender Verlauf
    ''' </summary>
    Private Function definecharacteristics(min As Double, max As Double, avg As Double) As Integer

        Dim bereich As Integer = CInt(UBound(_bedarf) / 4)
        Dim i As Integer
        Dim minvorne As Boolean = False, minhinten As Boolean = False, _
            maxvorne As Boolean = False, maxhinten As Boolean = False

        For i = 0 To bereich
            If _bedarf(i) = min Then
                minvorne = True
            ElseIf _bedarf(i) = max Then
                maxvorne = True
            End If
        Next i

        For i = UBound(_bedarf) - bereich To UBound(_bedarf)
            If _bedarf(i) = min Then
                minhinten = True
            ElseIf _bedarf(i) = max Then
                maxhinten = True
            End If
        Next

        If minvorne And maxhinten Then
            definecharacteristics = 1
        ElseIf maxvorne And minhinten Then
            definecharacteristics = 3
        Else
            definecharacteristics = 2
        End If

    End Function

    Public Sub New()
        ' erstmal keine Zuordnung zu einer TeamID  
        _teamID = -1
    End Sub

    ''' <summary>
    ''' laenge is actually dimension
    ''' </summary>
    ''' <param name="laenge"></param>
    Public Sub New(ByVal laenge As Integer)

        ReDim _bedarf(laenge)
        ' erstmal keine Zuordnung zu einer TeamID  
        _teamID = -1


    End Sub

End Class
