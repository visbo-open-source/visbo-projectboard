Imports System.Xml
Imports System.Xml.Schema


<Serializable()>
Public Class clsLicences

    Private _allLicenceKeys As SortedList(Of String, String)



    ''' <summary>
    ''' errechnet aus einem maximalen Datum, einer User Kennung und einer Komponenten Kennung den Schlüssel 
    ''' </summary>
    ''' <param name="untilDate"></param>
    ''' <param name="User"></param>
    ''' <param name="component"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property berechneKey(ByVal untilDate As Date, ByVal User As String, ByVal component As String) As String
        Get

            ''User = "Ute"
            ''component = "Swimmlanes2"
            ''untilDate = DateAdd(DateInterval.Month, 1200, Date.Now)         ' 100 Jahre Gültigkeit der Lizenz

            Dim userPrim As Long = 211

            ' User-Name in Kleinbuchstaben umwandeln
            User = LCase(User)

            ' Codierung Username
            Dim userlicCode As Long = 0
            Dim zahl(User.Length - 1) As Long

            For i As Integer = 0 To User.Length - 1
                zahl(i) = Convert.ToInt64(User(i))

                userlicCode = userlicCode + zahl(i) * userPrim
            Next
            Dim hexuserlicCode As String = Hex(userlicCode)

            ' Codierung Komponenten
            Dim compPrim As Long = 103

            Dim compLicCode As Long
            Dim hexcomplicCode As String


            compLicCode = 0

            Dim compzahl(component.Length - 1) As Long
            For i As Integer = 0 To component.Length - 1
                compzahl(i) = Convert.ToInt64(component(i))

                compLicCode = compLicCode + compzahl(i) * compPrim
            Next i

            compLicCode = compLicCode * userlicCode
            hexcomplicCode = Hex(compLicCode)


            ' Codierung validDate
            Dim validLicCode As Long = 0
            validLicCode = DateDiff(DateInterval.Day, Date.MinValue, untilDate)
            validLicCode = validLicCode * compLicCode

            Dim hexvalidLicCode As String = Hex(validLicCode)


            ' Call MsgBox("License: " & hexvalidLicCode & "-" & hexuserlicCode & "-" & hexcomplicCode)

            berechneKey = hexvalidLicCode & "-" & hexuserlicCode & "-" & hexcomplicCode

        End Get
    End Property

    ''' <summary>
    ''' checkt, ob ein gültiger Lizenz-KEy vorhanden ist
    ''' dazu werden alle ausgelesenen Lizenzkeys mit den Eingabe Werten user, komponente verglichen 
    ''' Es sollten folgende Meldungen kommen: 
    '''  
    ''' </summary>
    ''' <param name="user"></param>
    ''' <param name="component"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property validLicence(ByVal user As String, ByVal component As String) As Boolean

        Get
            Dim heute As Date = Date.Now
            Dim userPrim As Long = 211
            Dim compPrim As Long = 103

            Dim userfound = False
            Dim komponentefound = False

            validLicence = False

            ' User-Name in Kleinbuchstaben umwandeln
            user = LCase(user)


            ' '' '' Codierung Username
            '' ''Dim userlicCode As Long = 0
            '' ''Dim zahl(user.Length - 1) As Long

            '' ''For i As Integer = 0 To user.Length - 1
            '' ''    zahl(i) = Convert.ToInt64(user(i))

            '' ''    userlicCode = userlicCode + zahl(i) * userPrim
            '' ''Next
            '' ''Dim hexuserlicCode As String = Hex(userlicCode)


            ' '' '' Codierung Komponenten
            '' ''Dim compLicCode As Long
            '' ''Dim hexcomplicCode As String


            '' ''compLicCode = 0

            '' ''Dim compzahl(component.Length - 1) As Long
            '' ''For i As Integer = 0 To component.Length - 1
            '' ''    compzahl(i) = Convert.ToInt64(component(i))

            '' ''    compLicCode = compLicCode + compzahl(i) * compPrim
            '' ''Next i

            '' ''compLicCode = compLicCode * userlicCode
            '' ''hexcomplicCode = Hex(compLicCode)


            For Each kvp As KeyValuePair(Of String, String) In _allLicenceKeys

                Dim licStr() As String = Split(kvp.Key, "-", -1)

                Dim hexuserlic As String = licStr(1)
                Dim hexcomplic As String = licStr(2)
                Dim hexdatelic As String = licStr(0)

                Dim userlicCode As Long = Convert.ToInt64(hexuserlic, 16)

                ' User-LicensCode erzeugen für aktuellen User
                Dim hilfsUserlicCode As Long = 0

                Dim zahl(user.Length - 1) As Long

                For i As Integer = 0 To user.Length - 1
                    zahl(i) = Convert.ToInt64(user(i))

                    hilfsUserlicCode = hilfsUserlicCode + zahl(i) * userPrim
                Next

                Dim hexhilfsUlicCode As String = Hex(hilfsUserlicCode)
                If hexhilfsUlicCode = hexuserlic Then
                    '  Call MsgBox("user " & user & " vorhanden")
                End If

                ' User überprüfen
                If hilfsUserlicCode = userlicCode Then

                    userfound = True

                    ' die Komponente checken

                    Dim complicCode As Long = Convert.ToInt64(hexcomplic, 16)

                    ' Codierung Komponente "komponente"
                    Dim hilfscompLicCode As Long = 0

                    Dim compzahl(component.Length - 1) As Long
                    For i As Integer = 0 To component.Length - 1
                        compzahl(i) = Convert.ToInt64(component(i))

                        hilfscompLicCode = hilfscompLicCode + compzahl(i) * compPrim
                    Next i

                    hilfscompLicCode = hilfscompLicCode * userlicCode
                    Dim hexhilfsclicCode As String = Hex(hilfscompLicCode)

                    If hexhilfsclicCode = hexcomplic Then
                        '  Call MsgBox("komponente " & component & " für User: " & user & "vorhanden")
                    End If

                    If hilfscompLicCode = complicCode Then

                        komponentefound = True

                        ' Gültigkeitsdatum überprüfen

                        Dim datelicCode As Long = Convert.ToInt64(hexdatelic, 16)

                        ' Codierung validDate
                        Dim hilfsvalidLicCode As Long = 0
                        hilfsvalidLicCode = DateDiff(DateInterval.Day, Date.MinValue, heute)
                        hilfsvalidLicCode = hilfsvalidLicCode * complicCode

                        Dim hexhilfsvalidCode As String = Hex(hilfsvalidLicCode)
                        If hexhilfsvalidCode = hexdatelic Then
                            '   Call MsgBox("Date für User " & user & " und Komponente " & component & " ist ok")
                        End If

                        If hilfsvalidLicCode < datelicCode Then

                            validLicence = True

                        End If

                    End If

                End If

            Next

        End Get
    End Property
    Public ReadOnly Property Liste As SortedList(Of String, String)

        Get
            Liste = _allLicenceKeys
        End Get

    End Property
    Public Sub clear()
        _allLicenceKeys.Clear()
    End Sub

    Public Sub New()
        _allLicenceKeys = New SortedList(Of String, String)

    End Sub

End Class
