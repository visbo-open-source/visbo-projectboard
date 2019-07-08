Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports System.Windows
Imports System
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net.Http
Imports System.IO
Imports System.Drawing
Imports System.Globalization
Imports System.Web
Imports Microsoft.VisualBasic
Imports System.Security.Principal
Imports System.Net
Imports System.Text
Imports WebServerAcc
Imports MongoDbAccess
Public Class Request

    'public serverUriName ="http://visbo.myhome-server.de:3484" 
    'Public serverUriName As String = "http://localhost:3484"

    Private usedWebServer As Boolean = awinSettings.visboServer
    Private DBAcc As Object
    Private dbname As String
    Private dburl As String
    Private uname As String
    Private pwd As String


    ''' <summary>
    '''  'Verbindung mit der Datenbank aufbauen (mit Angabe von Username und Passwort)
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <param name="databaseName">entspricht beim Visbo-Rest-Server dem VisboCenter</param>
    ''' <param name="username"></param>
    ''' <param name="dbPasswort"></param>
    Public Function login(ByVal URL As String, ByVal databaseName As String, ByVal username As String, ByVal dbPasswort As String, ByRef err As clsErrorCodeMsg) As Boolean


        Dim loginOK As Boolean = False

        Try
            If usedWebServer Then

                Dim access As New WebServerAcc.Request
                loginOK = access.login(ServerURL:=URL, databaseName:=databaseName, username:=username, dbPasswort:=dbPasswort, err:=err)
                If loginOK Then
                    DBAcc = access
                    dbname = databaseName
                    dburl = URL
                    uname = username
                    pwd = dbPasswort
                Else
                    If err.errorCode = 407 Then   ' Proxy-Authentifizierung required
                        ' try is once more
                        loginOK = access.login(ServerURL:=URL, databaseName:=databaseName, username:=username, dbPasswort:=dbPasswort, err:=err)
                        If loginOK Then
                            DBAcc = access
                            dbname = databaseName
                            dburl = URL
                            uname = username
                            pwd = dbPasswort
                        End If
                    End If

                End If

            Else  'es wird eine MongoDB direkt adressiert

                Dim access As New MongoDbAccess.Request(databaseURL:=URL, databaseName:=databaseName, username:=username, dbPasswort:=dbPasswort)
                loginOK = access.createIndicesOnce()
                If loginOK Then
                    DBAcc = access
                End If
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in DBAccRequest-Login" & ex.Message)
        End Try

        login = loginOK

    End Function

    ''' <summary>
    ''' prüft die Verfügbarkeit der MongoDB bzw. ob ein Login bereits erfolgte, d.h. token vorhanden
    ''' </summary>
    ''' <returns></returns>
    Public Function pingMongoDb() As Boolean

        Dim result As Boolean = False
        Try
            If usedWebServer Then
                '' ur: 24.10.2018 nach Mail von MS: Test, da es eigentlich nichts nützt, kann trotzdem schief gehen
                ''Try
                ''    result = CType(DBAcc, WebServerAcc.Request).pingMongoDb()

                ''Catch ex As Exception

                ''    Dim hstr() As String = Split(ex.Message, ":")
                ''    If CInt(hstr(0)) = 401 Then
                ''        loginErfolgreich = login(dburl, dbname, uname, pwd)
                ''        If loginErfolgreich Then
                ''            result = CType(DBAcc, WebServerAcc.Request).pingMongoDb()
                ''        End If
                ''    Else
                ''        Throw New ArgumentException(ex.Message)
                ''    End If
                ''End Try

                result = True

            Else 'es wird eine MongoDB direkt adressiert

                result = CType(DBAcc, MongoDbAccess.Request).pingMongoDb()

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        pingMongoDb = result
    End Function


    ''' <summary>
    ''' Wenn ein ReST-Server adressiert wird, so kann mit dieser Routine über Email-aufforderung ein neues Passwort gesetzt werden.
    ''' </summary>
    ''' <returns></returns>
    Public Function pwforgotten(ByVal ServerURL As String, ByVal databaseName As String, ByVal username As String) As Boolean

        Dim result As Boolean = False
        Dim err As New clsErrorCodeMsg
        Try
            If usedWebServer Then

                Dim access As New WebServerAcc.Request
                result = access.pwforgotten(ServerURL:=ServerURL, databaseName:=databaseName, username:=username, err:=err)

            Else 'es wird eine MongoDB direkt adressiert, hier gibt es kein Password forgotten

                result = False

            End If

        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try

        pwforgotten = result
    End Function

    ''' <summary>
    ''' prüft ob der Projektname schon vorhanden ist (ggf. inkl. VariantName)
    ''' falls Variantname null ist oder leerer String wird nur der Projektname überprüft.
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <param name="variantname"></param>
    ''' <param name="storedAtorBefore"></param>
    ''' <returns></returns>
    Public Function projectNameAlreadyExists(ByVal projectname As String, ByVal variantname As String, ByVal storedAtorBefore As DateTime, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).projectNameAlreadyExists(projectname, variantname, storedAtorBefore, err)

                    If result = False Then

                        Select Case err.errorCode
                            Case 200 ' success
                                ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).projectNameAlreadyExists(projectname, variantname, storedAtorBefore, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).projectNameAlreadyExists(projectname, variantname, storedAtorBefore)
            End If

        Catch ex As Exception

        End Try

        projectNameAlreadyExists = result

    End Function


    ''' <summary>
    ''' bringt alle in der Datenbank vorkommenden TimeStamps zurück , in absteigender Sortierung
    ''' </summary>
    ''' <returns>Collection, absteigend sortiert</returns>
    Public Function retrieveZeitstempelFromDB() As Collection

        Dim resultCollection As New Collection

        Try

            If usedWebServer Then

                resultCollection.Add(StartofCalendar)

                '' ur: 2018.11.14: wird nur benötigt, um zu sehen, wann der erste Timestamp ist, über alle Projekte im VC betrachtet
                ''
                ''Try
                ''    resultCollection = CType(DBAcc, WebServerAcc.Request).retrieveZeitstempelFromDB()

                ''Catch ex As Exception


                ''    Dim hstr() As String = Split(ex.Message, ":")
                ''    If CInt(hstr(0)) = 401 Then                    ' Token is expired
                ''        loginErfolgreich = login(dburl, dbname, uname, pwd)
                ''        If loginErfolgreich Then
                ''            resultCollection = CType(DBAcc, WebServerAcc.Request).retrieveZeitstempelFromDB()
                ''        End If
                ''    Else
                ''        Throw New ArgumentException(ex.Message)
                ''    End If

                ''End Try


            Else 'es wird eine MongoDB direkt adressiert
                resultCollection = CType(DBAcc, MongoDbAccess.Request).retrieveZeitstempelFromDB()
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveZeitstempelFromDB = resultCollection

    End Function

    ''' <summary>
    ''' bringt für die angegebene Projekt-Variante alle Zeitstempel in absteigender Sortierung zurück 
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <returns>Collection, absteigend sortiert</returns>
    Public Function retrieveZeitstempelFirstLastFromDB(ByVal pvName As String, ByRef err As clsErrorCodeMsg) As Collection

        Dim ergebnisCollection As New Collection

        Try

            If usedWebServer Then
                Try
                    ergebnisCollection = CType(DBAcc, WebServerAcc.Request).retrieveZeitstempelFirstLastFromDB(pvName, err)

                    If ergebnisCollection.Count <= 0 Then

                        Select Case err.errorCode
                            Case 200 ' success
                                ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    ergebnisCollection = CType(DBAcc, WebServerAcc.Request).retrieveZeitstempelFirstLastFromDB(pvName, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else 'es wird eine MongoDB direkt adressiert

                Dim interResult As Collection = CType(DBAcc, MongoDbAccess.Request).retrieveZeitstempelFromDB(pvName)
                ' First TimeStamp
                ergebnisCollection.Add(interResult.Item(0))
                ' Last TimeStamp
                ergebnisCollection.Add(interResult.Item(interResult.Count - 1))

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveZeitstempelFirstLastFromDB = ergebnisCollection

    End Function



    ''' <summary>
    ''' bringt für die angegebene Projekt-Variante alle Zeitstempel in absteigender Sortierung zurück 
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <returns>Collection, absteigend sortiert</returns>
    Public Function retrieveZeitstempelFromDB(ByVal pvName As String, ByRef err As clsErrorCodeMsg) As Collection

        Dim ergebnisCollection As New Collection

        Try

            If usedWebServer Then
                Try

                    ergebnisCollection = CType(DBAcc, WebServerAcc.Request).retrieveZeitstempelFromDB(pvName, err)

                    If ergebnisCollection.Count <= 0 Then

                        Select Case err.errorCode
                            Case 200 ' success
                                ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    ergebnisCollection = CType(DBAcc, WebServerAcc.Request).retrieveZeitstempelFromDB(pvName, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                ergebnisCollection = CType(DBAcc, MongoDbAccess.Request).retrieveZeitstempelFromDB(pvName)
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveZeitstempelFromDB = ergebnisCollection

    End Function


    ''' <summary>
    '''  liest entweder alle Projekte im angegebenen Zeitraum 
    '''  oder aber alle Timestamps der übergebenen Projektvariante im angegeben Zeitfenster
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <param name="variantName"></param>
    ''' <param name="zeitraumStart"></param>
    ''' <param name="zeitraumEnde"></param>
    ''' <param name="storedEarliest"></param>
    ''' <param name="storedLatest"></param>
    ''' <param name="onlyLatest"></param>
    ''' <returns></returns>
    Public Function retrieveProjectsFromDB(ByVal vpid As String,
                                           ByVal projectname As String, ByVal variantName As String,
                                               ByVal zeitraumStart As DateTime, ByVal zeitraumEnde As DateTime,
                                               ByVal storedEarliest As DateTime, ByVal storedLatest As DateTime,
                                               ByVal onlyLatest As Boolean,
                                               ByRef err As clsErrorCodeMsg) _
                                               As SortedList(Of String, clsProjekt)

        Dim result As New SortedList(Of String, clsProjekt)

        Try

            If usedWebServer Then

                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectsFromDB(vpid, projectname, variantName,
                                                                                       zeitraumStart, zeitraumEnde,
                                                                                       storedEarliest, storedLatest, onlyLatest, err)

                    If result.Count <= 0 Then

                        Select Case err.errorCode
                            Case 200 ' success
                                ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectsFromDB(vpid, projectname, variantName,
                                                                                       zeitraumStart, zeitraumEnde,
                                                                                       storedEarliest, storedLatest, onlyLatest, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveProjectsFromDB(projectname, variantName, zeitraumStart, zeitraumEnde, storedEarliest, storedLatest, onlyLatest)
            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveProjectsFromDB = prepProjectsForRoles(result)

    End Function


    Public Function retrieveProjectNamesByPNRFromDB(ByVal projektKDNr As String, ByRef err As clsErrorCodeMsg) As Collection

        Dim result As New Collection

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectNamesByPNRFromDB(projektKDNr, err)

                    If result.Count <= 0 Then

                        Select Case err.errorCode
                            Case 200 ' success
                                ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectNamesByPNRFromDB(projektKDNr, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try



            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveProjectNamesByPNRFromDB(projektKDNr)
            End If


        Catch ex As Exception

        End Try

        retrieveProjectNamesByPNRFromDB = result

    End Function

    ''' <summary>
    ''' liest ein bestimmtes Projekt aus der DB (ggf. inkl. VariantName), das zum angegebenen Zeitpunkt das aktuelle war
    ''' falls Variantname null ist oder leerer String wird nur der Projektname überprüft.
    ''' </summary>
    '''  <param name="projectname"></param>
    ''' <param name="variantname"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveOneProjectfromDB(ByVal projectname As String,
                                             ByVal variantname As String,
                                             ByVal vpid As String,
                                             ByVal storedAtOrBefore As DateTime,
                                             ByRef err As clsErrorCodeMsg) As clsProjekt
        Dim result As clsProjekt = Nothing

        Try

            If usedWebServer Then

                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveOneProjectfromDB(projectname, variantname, vpid, storedAtOrBefore, err)

                    If IsNothing(result) Then

                        Select Case err.errorCode
                            ' tk 5.5. kann das hier überhaupt mit success rauskommen ? 
                            Case 200 ' success

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveOneProjectfromDB(projectname, variantname, vpid, storedAtOrBefore, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If
                Catch ex As Exception

                    Throw New ArgumentException(ex.Message)

                End Try

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveOneProjectfromDB(projectname, variantname, storedAtOrBefore)
            End If

        Catch ex As Exception

        End Try


        retrieveOneProjectfromDB = prepProjectForRoles(result)

    End Function



    ''' <summary>
    '''    benennt alle Projekte mit Namen oldName um 
    '''    aber nur, wenn der neue Name nicht schon in der Datenbank existiert 
    ''' </summary>
    ''' <param name="oldName"></param>
    ''' <param name="newName"></param>
    ''' <param name="userName"></param>
    ''' <returns>true : rename wurde durchgeführt
    '''          false: rename konnte nicht ausgeführt werden</returns>
    Public Function renameProjectsInDB(ByVal oldName As String, ByVal newName As String, ByVal userName As String, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Try


            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).renameProjectsInDB(oldName, newName, userName, err)

                    If result = False Then

                        Select Case err.errorCode
                            Case 200 ' success
                                ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).renameProjectsInDB(oldName, newName, userName, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).renameProjectsInDB(oldName, newName, userName)
            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        renameProjectsInDB = result

    End Function


    ''' <summary>
    ''' speichert ein einzelnes Projekt in der Datenbank
    ''' Zeitstempel wird aus den Projekt-Infos genommen
    ''' ein Protfolio Manager speicher immer mit entsprechendem Varianten-Name 
    ''' </summary>
    ''' <param name="projekt"></param>
    ''' <param name="userName"></param>
    ''' <returns></returns>
    Public Function storeProjectToDB(ByVal projekt As clsProjekt,
                                     ByVal userName As String,
                                     ByRef mergedProj As clsProjekt,
                                     ByRef err As clsErrorCodeMsg,
                                     Optional ByVal attrToStore As Boolean = False) As Boolean

        Dim result As Boolean = False
        Try

            ' tk 5.2.19 automatisch setzen des actualdatauntil, wenn denn der Modus AutoSet gesetzt ist 
            If awinSettings.autoSetActualDataDate Then
                projekt.actualDataUntil = projekt.timeStamp.AddDays(-1 * (projekt.timeStamp.Day + 1)).Date.AddHours(12)
            End If


            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).storeProjectToDB(projekt, userName, mergedProj, err, attrToStore)

                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).storeProjectToDB(projekt, userName, mergedProj, err, attrToStore)
                                End If

                            Case 409 ' Conflict

                                Call MsgBox(err.errorMsg)

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try



            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).storeProjectToDB(projekt, userName)
            End If

        Catch ex As Exception
            Throw New ArgumentException("storeProjectToDB:" & ex.Message)
        End Try

        storeProjectToDB = result
    End Function



    ''' <summary>
    ''' liefert alle Varianten Namen eines bestimmten Projektes zurück 
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <returns></returns>
    Public Function retrieveVariantNamesFromDB(ByVal projectName As String, ByRef err As clsErrorCodeMsg) As Collection

        Dim resultCollection As New Collection

        Try

            If usedWebServer Then
                Try
                    resultCollection = CType(DBAcc, WebServerAcc.Request).retrieveVariantNamesFromDB(projectName, err)

                    If resultCollection.Count <= 0 Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    resultCollection = CType(DBAcc, WebServerAcc.Request).retrieveVariantNamesFromDB(projectName, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                resultCollection = CType(DBAcc, MongoDbAccess.Request).retrieveVariantNamesFromDB(projectName)
            End If


        Catch ex As Exception
            Throw New ArgumentException("retrieveVariantNamesFromDB: " & ex.Message)
        End Try

        retrieveVariantNamesFromDB = resultCollection
    End Function


    ''' <summary>
    ''' liest alle vorkommenden Namen ProjektName#VariantenName aus der Datenbank , die zum Zeitpunkt storedLatest auch in der Datenbank existiert haben 
    ''' dabei wird ein übergebener Zeitraum berücksichtigt ... also nur Projekte, die auch im Zeitraum liegen ...
    ''' </summary>
    ''' <param name="zeitraumStart"></param>
    ''' <param name="zeitraumEnde"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveProjectVariantNamesFromDB(ByVal zeitraumStart As DateTime,
                                                          ByVal zeitraumEnde As DateTime,
                                                          ByVal storedAtOrBefore As DateTime,
                                                          ByRef err As clsErrorCodeMsg,
                                                          Optional ByVal fromReST As Boolean = False) _
                                                          As SortedList(Of String, String)

        Dim result As New SortedList(Of String, String)

        Try


            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectVariantNamesFromDB(zeitraumStart, zeitraumEnde, storedAtOrBefore, err, fromReST)

                    If result.Count <= 0 Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectVariantNamesFromDB(zeitraumStart, zeitraumEnde, storedAtOrBefore, err, fromReST)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveProjectVariantNamesFromDB(zeitraumStart, zeitraumEnde, storedAtOrBefore)
            End If

        Catch ex As Exception
            Throw New ArgumentException("retrieveProjectVariantNamesFromDB: " & ex.Message)
        End Try

        retrieveProjectVariantNamesFromDB = result

    End Function


    ''' <summary>
    ''' gibt die Projekthistorie innerhalb eines gegebenen Zeitraums zu einem gegebenen Projekt+Varianten-Namen zurück
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <param name="variantName"></param>
    ''' <param name="storedEarliest"></param>
    ''' <param name="storedLatest"></param>
    ''' <returns>sortierte Liste (DateTime, clsProjekt)</returns>
    Public Function retrieveProjectHistoryFromDB(ByVal projectname As String, ByVal variantName As String,
                                                 ByVal storedEarliest As DateTime, ByVal storedLatest As DateTime, ByRef err As clsErrorCodeMsg,
                                                 Optional ByVal vpid As String = "") As clsProjektHistorie

        Dim result As New clsProjektHistorie

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectHistoryFromDB(projectname, variantName,
                                                                                             storedEarliest, storedLatest, err, vpid)
                    If result.liste.Count <= 0 Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectHistoryFromDB(projectname, variantName,
                                                                                             storedEarliest, storedLatest, err, vpid)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try



            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveProjectHistoryFromDB(projectname, variantName, storedEarliest, storedLatest)
            End If

        Catch ex As Exception
            Throw New ArgumentException("retrieveProjectHistoryFromDB: " & ex.Message)
        End Try

        retrieveProjectHistoryFromDB = prepProjectsForRoles(result)

    End Function


    ''' <summary>
    ''' löscht den angegebenen TimeStamp der Projekt-Variante aus der Datenbank 
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <param name="variantName"></param>
    ''' <param name="stored"></param>
    ''' <param name="userName"></param>
    ''' <returns></returns>
    Public Function deleteProjectTimestampFromDB(ByVal projectname As String, ByVal variantName As String,
                                                     ByVal stored As DateTime, ByVal userName As String,
                                                     ByRef err As clsErrorCodeMsg) As Boolean
        Dim result As Boolean = False
        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).deleteProjectTimestampFromDB(projectname, variantName, stored, userName, err)

                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).deleteProjectTimestampFromDB(projectname, variantName, stored, userName, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).deleteProjectTimestampFromDB(projectname, variantName, stored, userName)
            End If

        Catch ex As Exception
            Throw New ArgumentException("deleteProjectTimestampFromDB: " & ex.Message)
        End Try

        If awinSettings.visboDebug Then
            Call MsgBox("Projekt mit timestamp: " & projectname & "/" & variantName & "/" & stored & "von User: " & userName & " gelöscht.")
        End If

        deleteProjectTimestampFromDB = result


    End Function


    ''' <summary>
    ''' holt die erste beauftragte Version des Projects 
    ''' immer mit Variant-Name = ""
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <returns></returns>
    Public Function retrieveFirstContractedPFromDB(ByVal projectname As String, ByVal variantname As String, ByRef err As clsErrorCodeMsg) As clsProjekt

        Dim result As New clsProjekt
        Try

            ' tk 28.12.18 
            ' es wird immer mit der durch den Portfolio Manager gemachten Vorgabe verglichen; und die hat immer den Varianten-Namen pfv (siehe Enum) 
            variantname = ptVariantFixNames.pfv.ToString

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveFirstContractedPFromDB(projectname, variantname, err)

                    If IsNothing(result) Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveFirstContractedPFromDB(projectname, variantname, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveFirstContractedPFromDB(projectname, variantname)
            End If

        Catch ex As Exception

            result = Nothing
            Throw New ArgumentException("retrieveFirstContractedPFromDB: " & ex.Message)
        End Try

        retrieveFirstContractedPFromDB = result

    End Function

    ''' <summary>
    ''' gibt den zum Zeitpunkt zuletzt beauftragten Stand zurück; bei Projekten muss variantNAme = "" sein, bei Summary Projekten VPortfolioName
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <param name="variantname"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveLastContractedPFromDB(ByVal projectname As String, ByVal variantname As String,
                                                  ByVal storedAtOrBefore As DateTime,
                                                  ByRef err As clsErrorCodeMsg) As clsProjekt

        Dim result As New clsProjekt

        Try
            ' tk 28.12.18 
            ' es wird immer mit der durch den Portfolio Manager gemachten Vorgabe verglichen; und die hat immer den Varianten-Namen pfv (siehe Enum) 
            variantname = ptVariantFixNames.pfv.ToString

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveLastContractedPFromDB(projectname, variantname, storedAtOrBefore, err)

                    If IsNothing(result) Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveLastContractedPFromDB(projectname, variantname, storedAtOrBefore, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).RetrieveLastContractedPFromDB(projectname, variantname, storedAtOrBefore)
            End If

        Catch ex As Exception

            result = Nothing
            Throw New ArgumentException("retrieveLastContractedPFromDB: " & ex.Message)
        End Try

        retrieveLastContractedPFromDB = result

    End Function



    ''' <summary>
    ''' überprüft, ob der User userName für das Projekt pvname vom Typ type 
    ''' die Erlaubnis hat etwas zu verändern
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <param name="userName"></param>
    ''' <param name="type"></param>
    ''' <returns>true -  es darf geändert werden
    '''          false - es darf nicht geändert werden</returns>
    Public Function checkChgPermission(ByVal pName As String,
                                       ByVal vName As String,
                                       ByVal userName As String,
                                       ByRef err As clsErrorCodeMsg,
                                       Optional type As Integer = 0) As Boolean

        Dim result As Boolean = False

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).checkChgPermission(pName, vName, userName, err, type)

                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).checkChgPermission(pName, vName, userName, err, type)
                                End If

                            Case Else ' all others
                                'Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).checkChgPermission(pName, vName, userName, type)
            End If

        Catch ex As Exception

            Throw New ArgumentException("checkChgPermission: " & ex.Message)
        End Try

        checkChgPermission = result

    End Function


    ''' <summary>
    ''' liefert für den pName und vName das clsWriteProtectiomItem zurück 
    ''' wenn es das nch nicht gibt, dann Null 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <param name="type"></param>
    ''' <returns></returns>
    Public Function getWriteProtection(ByVal pName As String, ByVal vName As String,
                                       ByRef err As clsErrorCodeMsg,
                                       Optional type As ptPRPFType = ptPRPFType.project) As clsWriteProtectionItem

        Dim result As New clsWriteProtectionItem

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).getWriteProtection(pName, vName, err, type)

                    If IsNothing(result) Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).getWriteProtection(pName, vName, err, type)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try



            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).getWriteProtection(pName, vName, type)
            End If

        Catch ex As Exception

            Throw New ArgumentException("getWriteProtection: " & ex.Message)
        End Try

        getWriteProtection = result

    End Function


    ''' <summary>
    ''' setzt für das entsprechende Item das Flag, dass es geschützt ist
    ''' gibt true zurück, wenn die Aktion erfolgreich war, false andernfalls
    ''' </summary>
    ''' <param name="wpItem"></param>
    ''' <returns></returns>
    Public Function setWriteProtection(ByVal wpItem As clsWriteProtectionItem,
                                       ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).setWriteProtection(wpItem, err)

                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).setWriteProtection(wpItem, err)
                                End If
                            Case 404    'DeleteVPLock: not Found
                                ' nothing to do

                            Case Else ' all others
                                'Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try



            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).setWriteProtection(wpItem)
            End If

        Catch ex As Exception

            Throw New ArgumentException("setWriteProtection: " & ex.Message)
        End Try

        setWriteProtection = result
    End Function

    ''' <summary>
    ''' Alle Projekte zu der Constellation 'portfolioName' zum Zeitpunkt storedAtOrBefore sortiert nach 'Projektname#VariantenName#'
    ''' </summary>
    ''' <param name="portfolioName"></param>
    ''' <param name="vpid"></param>
    ''' <param name="err"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveProjectsOfOneConstellationFromDB(ByVal portfolioName As String, ByVal vpid As String,
                                                             ByRef err As clsErrorCodeMsg,
                                                             Optional ByVal storedAtOrBefore As Date = Nothing) As SortedList(Of String, clsProjekt)

        Dim result As SortedList(Of String, clsProjekt) = Nothing

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectsOfOneConstellationFromDB(portfolioName, vpid, err, storedAtOrBefore)

                    If result.Count <= 0 Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveProjectsOfOneConstellationFromDB(portfolioName, vpid, err, storedAtOrBefore)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try



            Else 'es wird eine MongoDB direkt adressiert
                result = Nothing
            End If

        Catch ex As Exception

            Throw New ArgumentException("retrieveProjectsOfOneConstellationFromDB: " & ex.Message)
        End Try

        retrieveProjectsOfOneConstellationFromDB = result

    End Function

    ''' <summary>
    ''' liefert das Portfolios 'portfolioName' das zum storedAtOrBefore gespeichert war. In timestamp ist Datum/Uhrzeit dessen enthalten
    ''' </summary>
    ''' <param name="portfolioName"></param>
    ''' <param name="vpid"></param>
    ''' <param name="timestamp"></param>
    ''' <param name="err"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveOneConstellationFromDB(ByVal portfolioName As String, ByVal vpid As String,
                                                   ByRef timestamp As Date,
                                                   ByRef err As clsErrorCodeMsg,
                                                   Optional ByVal storedAtOrBefore As Date = Nothing) As clsConstellation

        Dim result As clsConstellation = Nothing

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveOneConstellationFromDB(portfolioName, vpid, timestamp,
                                                                                               err, storedAtOrBefore)

                    If Not IsNothing(result) Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveOneConstellationFromDB(portfolioName, vpid, timestamp,
                                                                                                             err, storedAtOrBefore)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else 'es wird eine MongoDB direkt adressiert
                result = Nothing
            End If

        Catch ex As Exception

            Throw New ArgumentException("retrieveOneConstellationFromDB: " & ex.Message)
        End Try

        retrieveOneConstellationFromDB = result

    End Function

    ''' <summary>
    ''' liefert das Portfolios 'portfolioName' das zum storedAtOrBefore gespeichert war. In timestamp ist Datum/Uhrzeit dessen enthalten
    ''' </summary>
    ''' <param name="portfolioName"></param>
    ''' <param name="err"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveFirstVersionOfOneConstellationFromDB(ByVal portfolioName As String, ByVal vpid As String,
                                                                 ByRef timestamp As Date,
                                                             ByRef err As clsErrorCodeMsg,
                                                             Optional ByVal storedAtOrBefore As Date = Nothing) As clsConstellation
        Dim result As clsConstellation = Nothing

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveFirstVersionOfOneConstellationFromDB(portfolioName, vpid, timestamp,
                                                                                                             err, storedAtOrBefore)

                    If Not IsNothing(result) Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveFirstVersionOfOneConstellationFromDB(portfolioName, vpid, timestamp,
                                                                                                             err, storedAtOrBefore)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else 'es wird eine MongoDB direkt adressiert
                result = Nothing
            End If

        Catch ex As Exception

            Throw New ArgumentException("retrieveFirstVersionOfOneConstellationFromDB: " & ex.Message)
        End Try

        retrieveFirstVersionOfOneConstellationFromDB = result

    End Function



    ''' <summary>
    '''  Alle Portfolios(Constellations) aus der Datenbank holen
    '''  Das Ergebnis dieser Funktion ist eine Liste (String, clsConstellation) 
    ''' </summary>
    ''' <returns></returns>
    Public Function retrieveConstellationsFromDB(ByVal storedAtOrBefore As Date, ByRef err As clsErrorCodeMsg) As clsConstellations

        Dim result As clsConstellations = Nothing

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveConstellationsFromDB(storedAtOrBefore, err)

                    If result.Count <= 0 Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveConstellationsFromDB(storedAtOrBefore, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try



            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveConstellationsFromDB()
            End If

        Catch ex As Exception

            Throw New ArgumentException("retrieveConstellationsFromDB: " & ex.Message)
        End Try

        retrieveConstellationsFromDB = result
    End Function


    ''' <summary>
    ''' Speichert ein Multiprojekt-Szenario in der Datenbank
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Public Function storeConstellationToDB(ByVal c As clsConstellation, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).storeConstellationToDB(c, err)
                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).storeConstellationToDB(c, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).storeConstellationToDB(c)
            End If

        Catch ex As Exception
            If awinSettings.visboDebug Then
                Throw New ArgumentException("storeConstellationToDB: " & ex.Message)
            Else
                Throw New ArgumentException(ex.Message)
            End If

        End Try

        storeConstellationToDB = result
    End Function

    ''' <summary>
    ''' Löschen des Portfolios  aus der Datenbank
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Public Function removeConstellationFromDB(ByVal c As clsConstellation, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).removeConstellationFromDB(c, err)

                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).removeConstellationFromDB(c, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).removeConstellationFromDB(c)
            End If

        Catch ex As Exception

            Throw New ArgumentException("removeConstellationFromDB: " & ex.Message)
        End Try

        removeConstellationFromDB = result
    End Function



    ''' <summary>
    '''  speichert einen Filter mit Namen 'name' in der Datenbank
    ''' </summary>
    ''' <param name="ptFilter"></param>
    ''' <param name="selfilter"></param>
    ''' <returns></returns>
    Public Function storeFilterToDB(ByVal ptFilter As clsFilter, ByRef selfilter As Boolean) As Boolean
        Dim result As Boolean = False

        Try

            If usedWebServer Then
                result = CType(DBAcc, WebServerAcc.Request).storeFilterToDB(ptFilter, selfilter)

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).storeFilterToDB(ptFilter, selfilter)
            End If

        Catch ex As Exception

            Throw New ArgumentException("storeFilterToDB: " & ex.Message)
        End Try

        storeFilterToDB = result
    End Function


    ''' <summary>
    ''' Alle Abhängigkeiten aus der Datenbank lesen
    ''' und als Ergebnis ein Liste von Abhängigkeiten zurückgeben
    ''' </summary>
    ''' <returns></returns>
    Public Function retrieveDependenciesFromDB() As clsDependencies

        Dim result As clsDependencies = Nothing
        Try

            If usedWebServer Then
                result = CType(DBAcc, WebServerAcc.Request).retrieveDependenciesFromDB()

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveDependenciesFromDB()
            End If

        Catch ex As Exception

            Throw New ArgumentException("retrieveDependenciesFromDB: " & ex.Message)
        End Try

        retrieveDependenciesFromDB = result
    End Function



    ''' <summary>
    ''' holt von allen Projekt-Varianten in AlleProjekte die Write-Protections
    ''' </summary>
    ''' <param name="AlleProjekte"></param>
    ''' <returns></returns>
    Public Function retrieveWriteProtectionsFromDB(ByVal AlleProjekte As clsProjekteAlle,
                                                   ByRef err As clsErrorCodeMsg) As SortedList(Of String, clsWriteProtectionItem)

        Dim result As New SortedList(Of String, clsWriteProtectionItem)
        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)

                    If result.Count <= 0 Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try



            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveWriteProtectionsFromDB(AlleProjekte)
            End If

        Catch ex As Exception

            Throw New ArgumentException("retrieveWriteProtectionsFromDB: " & ex.Message)
        End Try

        retrieveWriteProtectionsFromDB = result

    End Function


    ''' <summary>
    ''' löst von allen Projekt-Varianten des Users user die nonpermanent writeProtections
    ''' </summary>
    ''' <param name="user"></param>
    ''' <returns></returns>
    Public Function cancelWriteProtections(ByVal user As String, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).cancelWriteProtections(user, err)

                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).cancelWriteProtections(user, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).cancelWriteProtections(user)
            End If

        Catch ex As Exception

            Throw New ArgumentException("cancelWriteProtections: " & ex.Message)
        End Try

        cancelWriteProtections = result

    End Function


    ''' <summary>
    ''' liest alle Filter aus der Datenbank 
    ''' </summary>
    ''' <param name="selfilter"></param>
    ''' <returns></returns>
    Public Function retrieveAllFilterFromDB(ByVal selfilter As Boolean) As SortedList(Of String, clsFilter)
        Dim result As New SortedList(Of String, clsFilter)
        Try

            If usedWebServer Then
                result = CType(DBAcc, WebServerAcc.Request).retrieveAllFilterFromDB(selfilter)

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveAllFilterFromDB(selfilter)
            End If

        Catch ex As Exception

            Throw New ArgumentException("retrieveAllFilterFromDB: " & ex.Message)
        End Try

        retrieveAllFilterFromDB = result
    End Function


    ''' <summary>
    ''' löscht einen bestimmten Filter aus der Datenbank
    ''' </summary>
    ''' <param name="filter"></param>
    ''' <returns></returns>
    Public Function removeFilterFromDB(ByVal filter As clsFilter) As Boolean
        Dim result As Boolean = False
        Try

            If usedWebServer Then
                result = CType(DBAcc, WebServerAcc.Request).removeFilterFromDB(filter)

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).removeFilterFromDB(filter)
            End If

        Catch ex As Exception

            Throw New ArgumentException("removeFilterFromDB: " & ex.Message)
        End Try

        removeFilterFromDB = result

    End Function

    ''' <summary>
    ''' liest die Rollendefinitionen aus der Datenbank aus Collection vcrole
    ''' </summary>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveRolesFromDB(ByVal storedAtOrBefore As DateTime, ByRef err As clsErrorCodeMsg) As clsRollen

        Dim result As New clsRollen()

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveRolesFromDB(storedAtOrBefore, err)

                    If result.Count <= 0 Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveRolesFromDB(storedAtOrBefore, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveRolesFromDB(storedAtOrBefore)
            End If

        Catch ex As Exception

            Throw New ArgumentException("retrieveRolesFromDB: " & ex.Message)
        End Try

        retrieveRolesFromDB = result

    End Function



    ''' <summary>
    '''  liest die Kostenartdefinitionen aus der Datenbank aus Collection vccost
    ''' </summary>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveCostsFromDB(ByVal storedAtOrBefore As DateTime, ByRef err As clsErrorCodeMsg) As clsKostenarten

        Dim result As New clsKostenarten()
        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveCostsFromDB(storedAtOrBefore, err)
                    If result.Count <= 0 Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveCostsFromDB(storedAtOrBefore, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try



            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).retrieveCostsFromDB(storedAtOrBefore)
            End If

        Catch ex As Exception

            Throw New ArgumentException("retrieveCostsFromDB: " & ex.Message)
        End Try

        retrieveCostsFromDB = result

    End Function


    ''' <summary>
    ''' speichert eine Rolle in der Datenbank; 
    ''' wenn insertNewDate = true: speichere eine neue Timestamp-Instanz 
    ''' andernfalls wird die Rolle Replaced 
    ''' </summary>
    ''' <param name="role"></param>
    ''' <param name="insertNewDate"></param>
    ''' <param name="ts"></param>
    ''' <returns></returns>
    Public Function storeRoleDefinitionToDB(ByVal role As clsRollenDefinition, ByVal insertNewDate As Boolean, ByVal ts As DateTime, ByRef err As clsErrorCodeMsg) As Boolean
        Dim result As Boolean = False

        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).storeRoleDefinitionToDB(role, insertNewDate, ts, err)

                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).storeRoleDefinitionToDB(role, insertNewDate, ts, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).storeRoleDefinitionToDB(role, insertNewDate, ts)
            End If

        Catch ex As Exception

            Throw New ArgumentException("storeRoleDefinitionToDB: " & ex.Message)
        End Try

        storeRoleDefinitionToDB = result
    End Function
    ''' <summary>
    '''  speichert eine Kostenart In der Datenbank; 
    '''  wenn insertNewDate = True: speichere eine neue Timestamp-Instanz 
    '''  andernfalls wird die Kostenart Replaced, sofern sie sich geändert hat  
    ''' </summary>
    ''' <param name="cost"></param>
    ''' <param name="insertNewDate"></param>
    ''' <param name="ts"></param>
    ''' <returns></returns>
    Public Function storeCostDefinitionToDB(ByVal cost As clsKostenartDefinition, ByVal insertNewDate As Boolean, ByVal ts As DateTime, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).storeCostDefinitionToDB(cost, insertNewDate, ts, err)
                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).storeCostDefinitionToDB(cost, insertNewDate, ts, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try


            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).storeCostDefinitionToDB(cost, insertNewDate, ts)
            End If

        Catch ex As Exception

            Throw New ArgumentException("storeCostDefinitionToDB: " & ex.Message)
        End Try
        storeCostDefinitionToDB = result

    End Function

    ''' <summary>
    ''' speichert Projekt-Dependencies in DB 
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function storeDependencyofPToDB(ByVal d As clsDependenciesOfP) As Boolean

        Dim result As Boolean = False

        Try

            If usedWebServer Then
                result = CType(DBAcc, WebServerAcc.Request).storeDependencyofPToDB(d)

            Else 'es wird eine MongoDB direkt adressiert
                result = CType(DBAcc, MongoDbAccess.Request).storeDependencyofPToDB(d)
            End If

        Catch ex As Exception

            Throw New ArgumentException("storeDependencyofPToDB: " & ex.Message)
        End Try
        storeDependencyofPToDB = result

    End Function


    ''' <summary>
    ''' speichert VC Settings in DB 
    ''' </summary>
    ''' <param name="hlist"></param>
    ''' <param name="type"></param>
    ''' <param name="ts"></param>
    ''' <param name="err"></param>
    ''' <returns></returns>
    Public Function storeVCSettingsToDB(ByVal hlist As Object,
                                        ByVal type As String,
                                        ByVal name As String,
                                        ByVal ts As Date,
                                        ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try
            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).storeVCsettingsToDB(hlist, type, name, ts, err)

                    If result = False Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).storeVCsettingsToDB(hlist, type, name, ts, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else 'es wird eine MongoDB direkt adressiert; hier gibt es keine Settings

                result = False
            End If

        Catch ex As Exception

            'Call MsgBox("storeVCSettingsToDB: " & ex.Message)
        End Try
        storeVCSettingsToDB = result

    End Function
    Public Function retrieveCustomUserRoles(ByRef err As clsErrorCodeMsg) As clsCustomUserRoles

        Dim result As New clsCustomUserRoles

        Try
            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveCustomUserRoles(err)

                    If Not IsNothing(result) Then

                        If result.count <= 0 Then

                            Select Case err.errorCode

                                Case 200 ' success
                                     ' nothing to do

                                Case 401 ' Token is expired
                                    loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                    If loginErfolgreich Then
                                        result = CType(DBAcc, WebServerAcc.Request).retrieveCustomUserRoles(err)
                                    End If

                                Case Else ' all others
                                    Throw New ArgumentException(err.errorMsg)
                            End Select

                        End If

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else
                ' nothing can be done for direct MongoAccess
            End If

        Catch ex As Exception

        End Try
        retrieveCustomUserRoles = result
    End Function
    Public Function retrieveOrganisationFromDB(ByVal name As String,
                                          ByVal timestamp As Date,
                                          ByVal refnext As Boolean,
                                          ByRef err As clsErrorCodeMsg) As clsOrganisation

        Dim result As New clsOrganisation

        Try
            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveOrganisationFromDB("", timestamp, refnext, err)

                    If Not IsNothing(result) Then
                        If result.count <= 0 Then

                            Select Case err.errorCode

                                Case 200 ' success
                                     ' nothing to do

                                Case 401 ' Token is expired
                                    loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                    If loginErfolgreich Then
                                        result = CType(DBAcc, WebServerAcc.Request).retrieveOrganisationFromDB("", timestamp, refnext, err)
                                    End If

                                Case Else ' all others
                                    Throw New ArgumentException(err.errorMsg)
                            End Select

                        End If
                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else
                ' to do for direct MongoAccess
                result.allRoles = CType(DBAcc, MongoDbAccess.Request).retrieveRolesFromDB(timestamp)
                result.allCosts = CType(DBAcc, MongoDbAccess.Request).retrieveCostsFromDB(timestamp)
                result.validFrom = StartofCalendar

            End If

        Catch ex As Exception

        End Try

        retrieveOrganisationFromDB = result
    End Function


    Public Function retrieveCustomFieldsFromDB(ByRef err As clsErrorCodeMsg) As clsCustomFieldDefinitions

        Dim result As New clsCustomFieldDefinitions

        Try
            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveCustomFieldsFromDB(err)

                    If Not IsNothing(result) Then

                        If result.count <= 0 Then

                            Select Case err.errorCode

                                Case 200 ' success
                                     ' nothing to do

                                Case 401 ' Token is expired
                                    loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                    If loginErfolgreich Then
                                        result = CType(DBAcc, WebServerAcc.Request).retrieveCustomFieldsFromDB(err)
                                    End If

                                Case Else ' all others
                                    Throw New ArgumentException(err.errorMsg)
                            End Select

                        End If

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try

            Else
                ' to do for direct MongoAccess
                result = Nothing
                err.errorCode = 403
                err.errorMsg = "Fehler: CustomFields sind im Customization-File gespeichert " &
                                vbLf & "und können daher nicht von der DB gelesen werden"

            End If

        Catch ex As Exception

        End Try

        retrieveCustomFieldsFromDB = result
    End Function
    Public Function retrieveUserIDFromName(ByVal username As String, ByRef err As clsErrorCodeMsg) As String

        Dim result As String = ""
        Try

            If usedWebServer Then
                Try
                    result = CType(DBAcc, WebServerAcc.Request).retrieveUserIDFromName(dbUsername, err)

                    If result.Count <= 0 Then

                        Select Case err.errorCode

                            Case 200 ' success
                                     ' nothing to do

                            Case 401 ' Token is expired
                                loginErfolgreich = login(dburl, dbname, uname, pwd, err)
                                If loginErfolgreich Then
                                    result = CType(DBAcc, WebServerAcc.Request).retrieveUserIDFromName(dbUsername, err)
                                End If

                            Case Else ' all others
                                Throw New ArgumentException(err.errorMsg)
                        End Select

                    End If

                Catch ex As Exception
                    Throw New ArgumentException(ex.Message)
                End Try
            Else
                ' nothing to do for direct MongoAccess
            End If

        Catch ex As Exception

        End Try
        retrieveUserIDFromName = result
    End Function

    Public Function clearCache()
        Dim result As Boolean = False
        Try
            result = CType(DBAcc, WebServerAcc.Request).clearVRSCache()
        Catch ex As Exception
            Call MsgBox("Fehler beim löschen des Cache")
        End Try
        clearCache = result
    End Function
End Class
