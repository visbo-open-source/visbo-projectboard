
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
Public Class Request

    'public serverUriName ="http://visbo.myhome-server.de:3484" 
    Public serverUriName As String = "http://localhost:3484"

    Public aktVCid As String = ""

    Private token As String = ""
    Private VCs As New List(Of clsVC)
    Private VPs As New List(Of clsVP)
    Private aktUser As clsUserReg = Nothing

    Private webVCs As clsWebVC = Nothing

    Private aktVC As clsWebVC = Nothing
    Private webVPs As clsWebVP = Nothing

    Private aktVP As clsWebVP = Nothing
    Private webVPvs As clsWebVPv = Nothing
    Private aktVPv As clsWebLongVPv = Nothing




    ''' <summary>
    ''' Sendet einen Request vom Typ method an den Server. Außerdem wird hier auch die Antwort empfangen und an die aufrufenden Routine zurückgegeben
    ''' </summary>
    ''' <param name="uri">Url fur den REst-Request</param>
    ''' <param name="data">Daten für die Aufrufe von POST/PUT</param>
    ''' <param name="method">Typ des Rest-Request  GET/POST/PUT/DELETE</param>
    Private Function GetRestServerResponse(ByVal uri As Uri, ByVal data As Byte(), ByVal method As String) As HttpWebResponse
        'Private Function GetRestServerResponse(ByVal uri As Uri, ByVal data As Byte(), ByVal method As String) As HttpWebResponse

        Dim response As HttpWebResponse = Nothing

        Try
            Dim request As HttpWebRequest = DirectCast(HttpWebRequest.Create(uri), HttpWebRequest)

            request.Method = method
            request.ContentType = "application/json"
            request.Headers.Add("access-key", token)
            request.UserAgent = "VISBO Browser/x.x (" & My.Computer.Info.OSFullName & ":" & My.Computer.Info.OSPlatform & ":" & My.Computer.Info.OSVersion & ") Client:VISBO Projectboard/3.5 "

            request.ContentLength = data.Length

            If request.ContentLength > 0 Then
                Try
                    Using requestStream As Stream = request.GetRequestStream()
                        ' Send the data.
                        requestStream.Write(data, 0, data.Length)
                        requestStream.Close()
                        requestStream.Dispose()
                    End Using
                Catch ex As Exception
                    'Call MsgBox("Fehler bei GetRequestStream:  " & ex.Message)
                    Throw New ArgumentException("Fehler bei GetRequestStream:  " & ex.Message)
                End Try
            End If

            Try
                response = request.GetResponse()

            Catch ex As WebException

                response = ex.Response
            End Try

        Catch ex1 As Exception
            Call MsgBox(ex1.Message)
            Throw
        End Try

        Return response

    End Function

    Private Function ReadResponseContent(ByRef httpresp As HttpWebResponse) As String
        'Private Function ReadResponseContent(ByRef resp As HttpWebResponse) As String
        Try

            If IsNothing(httpresp) Then
                Throw New ArgumentNullException("HttpWebResponse ist Nothing")
            Else
                Dim statcode As HttpStatusCode = httpresp.StatusCode
                If statcode <> HttpStatusCode.OK Then
                    Call MsgBox(statcode.ToString & ":" & httpresp.StatusDescription)
                    Throw New ArgumentException(statcode.ToString & ":" & httpresp.StatusDescription)
                Else
                    Using sr As New StreamReader(httpresp.GetResponseStream)
                        Return sr.ReadToEnd()
                    End Using
                End If
            End If

        Catch ex As Exception
            Throw New ArgumentException("ReadResponseContent:" & ex.Message)
        End Try
    End Function


    ''' <summary>
    ''' diese Funktion konvertiert die Struktur, die für diesen Server-Request benötigt wird (type) in ein ByteArray im Json-Format
    ''' </summary>
    ''' <param name="dataClass"></param>
    ''' <param name="type"></param>
    ''' <returns>Object</returns>
    Private Function serverInputDataJson(ByVal dataClass As Object, ByVal type As String) As Byte()
        'Private Function serverInputDataJson(ByVal dataClass As Object, ByVal type As String) As Byte()

        serverInputDataJson = Nothing
        Dim encoding As New System.Text.UTF8Encoding()
        Dim bytes() As Byte = Nothing
        'Dim bufferlge As Int32 = 256
        'Dim ms As New MemoryStream(bufferlge)
        Dim hstr As String = ""
        'Dim ok As Boolean = True

        Try
            hstr = JsonConvert.SerializeObject(dataClass)
            'serverInputDataJson = encoding.GetBytes(hstr)
            serverInputDataJson = encoding.GetBytes(JsonConvert.SerializeObject(dataClass))

        Catch ex As Exception
            Call MsgBox("Fehler in serverInputDataJson " & type & ": " & ex.Message)
        End Try

    End Function

    ''' <summary>
    '''  'Verbindung mit der Datenbank aufbauen (mit Angabe von Username und Passwort)
    ''' </summary>
    ''' <param name="ServerURL"></param>
    ''' <param name="databaseName">wird beim Login am Visbo-Rest-Server nicht benötigt</param>
    ''' <param name="username"></param>
    ''' <param name="dbPasswort"></param>
    Public Function login(ByVal ServerURL As String, ByVal databaseName As String, ByVal username As String, ByVal dbPasswort As String) As String

        Dim typeRequest As String = "/token/user/login"
        'Dim typeRequest As String = "/token/user/signup"
        Dim serverUri As New Uri(ServerURL & typeRequest)
        Dim loginOK As Boolean = False

        Try

            Dim user As New clsUserLoginSignup
            user.email = username
            user.password = dbPasswort
            'user.email = "markus.seyfried@visbo.de"
            'user.password = "visbo123"

            ' Konvertiere die erforderlichen Inputdaten des Requests vom Typ typeRequest (von der Struktur cls??) in ein Json-ByteArray
            Dim data() As Byte
            data = serverInputDataJson(user, typeRequest)


            Dim loginAntwort As clsWebTokenUserLoginSignup
            Dim Antwort As String
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                loginAntwort = JsonConvert.DeserializeObject(Of clsWebTokenUserLoginSignup)(Antwort)
            End Using

            Call MsgBox(loginAntwort.message)
            loginOK = (loginAntwort.state = "success")

            If loginOK Then
                token = loginAntwort.token
                serverUriName = ServerURL
                aktUser = loginAntwort.user
                ' VisboCenterID mit Name = databaseName wird gespeichert
                aktVCid = GETvcid(databaseName)

            Else
                token = ""
                serverUriName = ServerURL
                aktUser = Nothing
            End If


        Catch ex As Exception
            Throw New ArgumentException("Fehler in PTWebRequestLogin" & typeRequest & ": " & ex.Message)
        End Try

        login = loginOK

    End Function

    ''' <summary>
    ''' prüft die Verfügbarkeit der MongoDB bzw. ob ein Login bereits erfolgte, d.h. token vorhanden
    ''' </summary>
    ''' <returns></returns>
    Public Function pingMongoDb() As Boolean

        Dim result As Boolean = False
        If token <> "" Then
            result = True
        End If

        pingMongoDb = result
    End Function

    ''' <summary>
    ''' prüft ob der Projektname schon vorhanden ist (ggf. inkl. VariantName)
    ''' falls Variantname null ist oder leerer String wird nur der Projektname überprüft.
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <param name="variantname"></param>
    ''' <param name="storedAtorBefore"></param>
    ''' <returns></returns>
    Public Function projectNameAlreadyExists(ByVal projectname As String, ByVal variantname As String, ByVal storedAtorBefore As DateTime) As Boolean

        Dim result As Boolean = False

        Try
            Dim vpid As String = ""
            vpid = GETvpid(projectname)

            If vpid <> "" Then
                ' gewünschte Variante vom Server anfordern
                Dim allVPv As New List(Of clsProjektWebShort)
                allVPv = GETallVPvShort(vpid, variantname, storedAtorBefore)
                result = (allVPv.Count > 0)
            End If

        Catch ex As Exception

        End Try

        projectNameAlreadyExists = result

    End Function


    ''' <summary>
    ''' bringt alle in der Datenbank vorkommenden TimeStamps zurück , in absteigender Sortierung
    ''' </summary>
    ''' <returns></returns>
    Public Function retrieveZeitstempelFromDB() As Collection

        Dim resultCollection As New Collection

        Try

            ' alle VisboProjectVersions vom Server anfordern
            Dim allVPv As New List(Of clsProjektWebShort)
            allVPv = GETallVPvShort("")

            ' alle vorhandenen Timestamps in der resultCollection sammeln
            Dim sl As New SortedList(Of Date, Date)
            For Each shortproj As clsProjektWebShort In allVPv
                If Not sl.ContainsKey(shortproj.timestamp) Then
                    sl.Add(shortproj.timestamp, shortproj.timestamp)
                End If
            Next
            For Each kvp As KeyValuePair(Of DateTime, DateTime) In sl
                resultCollection.Add(kvp.Value)
            Next

        Catch ex As Exception

        End Try

        retrieveZeitstempelFromDB = resultCollection

    End Function

    ''' <summary>
    ''' bringt für die angegebene Projekt-Variante alle Zeitstempel in absteigender Sortierung zurück 
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <returns></returns>
    Public Function retrieveZeitstempelFromDB(ByVal pvName As String) As Collection

        Dim ergebnisCollection As New Collection

        Try

            Dim projectName As String = ""
            Dim variantName As String = ""
            Dim vpid As String = ""

            Dim hstr() As String = Split(pvName, "#")
            If hstr.Length > 0 Then
                projectName = hstr(0)
            End If
            If hstr.Length > 1 Then
                variantName = hstr(1)
            End If

            ' VPID zu Projekt projectName holen vom WebServer/DB
            vpid = GETvpid(projectName)

            If vpid <> "" Then
                ' gewünschte Variante vom Server anfordern
                Dim allVPv As New List(Of clsProjektWebShort)
                allVPv = GETallVPvShort(vpid, variantName)

                ' alle vorhandenen Timestamps zu einem pvName in die ErgebnisCollection sammeln
                Dim sl As New SortedList(Of Date, Date)
                For Each shortproj As clsProjektWebShort In allVPv
                    If Not sl.ContainsKey(shortproj.timestamp) Then
                        sl.Add(shortproj.timestamp, shortproj.timestamp)
                    End If
                Next
                For Each kvp As KeyValuePair(Of DateTime, DateTime) In sl
                    ergebnisCollection.add(kvp.value)
                Next

            End If

        Catch ex As Exception

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
    Public Function retrieveProjectsFromDB(ByVal projectname As String, ByVal variantName As String,
                                               ByVal zeitraumStart As DateTime, ByVal zeitraumEnde As DateTime,
                                               ByVal storedEarliest As DateTime, ByVal storedLatest As DateTime,
                                               ByVal onlyLatest As Boolean) _
                                               As SortedList(Of String, clsProjekt)

        Dim result As New SortedList(Of String, clsProjekt)

        Try
            Dim hproj As New clsProjekt

            ' da in der Datenbank alle DateTime im UTC gespeichert sind, muss hier auch dieses Format verwendet werden
            storedLatest = storedLatest.ToUniversalTime()
            storedEarliest = storedEarliest.ToUniversalTime()

            ' Kein Projekt  angegeben. es werden alle Projekte im angebenen Zeitraum zurückgegeben

            If projectname = "" Then


                VPs = GETallVP(aktVCid)

                ' schleife über alle VisboProjects
                For Each vp As clsVP In VPs

                    Dim vpid As String = vp._id

                    If vpid <> "" Then
                        ' gewünschten Varianten vom Server anfordern
                        Dim allVPv As New List(Of clsProjektWebLong)
                        allVPv = GETallVPvLong(vpid, , variantName)

                        For Each webProj As clsProjektWebLong In allVPv

                            If (webProj.startDate <= zeitraumEnde And
                                webProj.endDate >= zeitraumStart And
                                webProj.timestamp <= storedLatest) Then

                                webProj.copyto(hproj)
                                Dim a As Integer = hproj.dauerInDays
                                Dim key As String = Projekte.calcProjektKey(hproj)
                                If Not result.ContainsKey(key) Then
                                    result.Add(key, hproj)
                                End If

                            End If

                        Next
                    Else
                        ' kann eigentlich nicht vorkommen
                    End If

                Next

            Else
                '  Projekt angegeben: d.h. es werden alle Timestamps der übergebenen Projekt-Variante zurückgegeben
                Dim vpid As String = GETvpid(projectname)
                If vpid <> "" Then
                    ' gewünschten Varianten vom Server anfordern
                    Dim allVPv As New List(Of clsProjektWebLong)
                    allVPv = GETallVPvLong(vpid, , variantName, storedLatest)

                    For Each webProj As clsProjektWebLong In allVPv
                        If webProj.timestamp >= storedEarliest Then

                            webProj.copyto(hproj)
                            Dim a As Integer = hproj.dauerInDays
                            Dim key As String = Projekte.calcProjektKey(hproj)
                            If Not result.ContainsKey(key) Then
                                result.Add(key, hproj)
                            End If

                        End If

                    Next

                End If

            End If


        Catch ex As Exception

        End Try

        retrieveProjectsFromDB = result

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
                                             ByVal storedAtOrBefore As DateTime) As clsProjekt
        Dim result As clsProjekt = Nothing

        Try
            Dim hproj As New clsProjekt
            Dim vpid As String = ""
            vpid = GETvpid(projectname)

            If vpid <> "" Then
                ' gewünschte Variante vom Server anfordern
                Dim allVPv As New List(Of clsProjektWebLong)
                allVPv = GETallVPvLong(vpid, variantname, storedAtOrBefore)
                If allVPv.Count = 1 Then
                    Dim webProj As clsProjektWebLong = allVPv.ElementAt(0)
                    webProj.copyto(hproj)
                    result = hproj
                End If

            End If

        Catch ex As Exception

        End Try
        retrieveOneProjectfromDB = result

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
    Public Function renameProjectsInDB(ByVal oldName As String, ByVal newName As String, ByVal userName As String) As Boolean

        Dim result As Boolean = False
        Try
            If projectNameAlreadyExists(newName, "", DateTime.Now) Then

                renameProjectsInDB = result

            Else

                Dim chkOk As Boolean = True

                ' hier wird überprüft, ob das Projekt selbst
                ' und auch keine der Varianten von einem anderen User schreibgeschützt ist

                chkOk = checkChgPermission(oldName, "", userName)

                If chkOk Then

                    Dim vpid As String = GETvpid(oldName)
                    If vpid <> "" Then
                        For Each vp As clsVP In VPs
                            If vp._id = vpid And vp.name = oldName Then
                                vp.name = newName
                            End If
                        Next
                    End If
                    Dim vpList As List(Of clsVP) = PUTOneVP(vpid)
                    ' rename war korrekt, wenn in vplist ein und zwar nur ein VisboProject zurückgegeben wurde.
                    result = (vpList.Count = 1)

                End If

            End If
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        renameProjectsInDB = result

    End Function



    ''' <summary>
    ''' speichert ein einzelnes Projekt in der Datenbank
    ''' Zeitstempel wird aus den Projekt-Infos genommen
    ''' </summary>
    ''' <param name="projekt"></param>
    ''' <param name="userName"></param>
    ''' <returns></returns>
    Public Function storeProjectToDB(ByVal projekt As clsProjekt, ByVal userName As String) As Boolean

        Dim result As Boolean = False
        Try
            Dim typeRequest As String = "/vpv"
            Dim serverUriString As String = serverUriName & typeRequest
            Dim serverUri As New Uri(serverUriString)
            Dim webVP As New clsWebVP
            Dim data() As Byte

            Dim pname As String = ""
            Dim vname As String = ""

            Dim hstr() As String = Split(projekt.name, "#")
            If hstr.Length > 0 Then
                ' projektName steht im ersten Teil
                pname = hstr(0)
            End If
            If hstr.Length > 1 Then
                ' variantName steht im zeiten Teil
                vname = hstr(1)
            End If

            Dim vpid As String = GETvpid(pname)
            Dim storedVP As Boolean = (vpid <> "")

            If Not storedVP Then
                Dim VP As New clsVP
                Dim user As New clsUser
                user.email = aktUser.email
                user.role = "Admin"
                VP.users.Add(user)
                VP.name = pname
                VP.vcid = aktVCid
                data = serverInputDataJson(VP, typeRequest)

                Dim Antwort As String
                Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                    Antwort = ReadResponseContent(httpresp)
                    webVP = JsonConvert.DeserializeObject(Of clsWebVP)(Antwort)
                End Using


                Call MsgBox(webVP.message)

                If webVP.state = "success" Then
                    ' vpid für neues Projekt merken, wird für speichern von vpv benötigt
                    vpid = webVP.vp.ElementAt(0)._id
                    storedVP = (vpid <> "")
                End If

            End If

            ' Projekt ist bereits in VisboProjects Collection gespeichert, es existiert eine vpid
            If storedVP Then

                If checkChgPermission(pname, vname, userName) Then

                    Dim projektWeb As New clsProjektWebLong
                    projektWeb.copyfrom(projekt)
                    projektWeb.origId = projektWeb.name & "#" & projektWeb.variantName & "#" & projektWeb.timestamp.ToString()
                    projektWeb.vpid = vpid
                    data = serverInputDataJson(projektWeb, "")

                    Dim storeAntwort As clsWebLongVPv
                    Dim Antwort As String
                    Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                        Antwort = ReadResponseContent(httpresp)
                        storeAntwort = JsonConvert.DeserializeObject(Of clsWebLongVPv)(Antwort)
                    End Using


                    Call MsgBox(storeAntwort.message)
                    result = (storeAntwort.state = "success")

                End If
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
    Public Function retrieveVariantNamesFromDB(ByVal projectName As String) As Collection

        Dim ergebnisCollection As New Collection

        Try
            Dim vpid As String = ""

            ' VPID zu Projekt projectName holen vom WebServer/DB
            vpid = GETvpid(projectName)

            If vpid <> "" Then
                ' von allen Varianten des Projektes vpid die neueste Version holen
                Dim allVPv As New List(Of clsProjektWebShort)
                allVPv = GETallVPvShort(vpid, storedAtorBefore:=Date.Now)

                ' alle Variantenamen in der Collection sammeln
                For Each shortproj As clsProjektWebShort In allVPv
                    ergebnisCollection.Add(shortproj.variantName, shortproj.variantName)
                Next
            End If

        Catch ex As Exception

        End Try

        retrieveVariantNamesFromDB = ergebnisCollection
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
                                                          ByVal storedAtOrBefore As DateTime) _
                                                          As SortedList(Of String, String)

        Dim result As New SortedList(Of String, String)

        Try
            ' Datum in der Datenbank ist UTC
            storedAtOrBefore = storedAtOrBefore.ToUniversalTime()

            ' holt alle Projekte/Variante/versionen mit ReferenzDatum storedatOrBefore
            Dim vpvListe As New List(Of clsProjektWebShort)
            vpvListe = GETallVPvShort("", "", storedAtOrBefore)

            For Each vpv As clsProjektWebShort In vpvListe

                If vpv.startDate <= zeitraumEnde And
                   vpv.endDate >= zeitraumStart Then

                    Dim pName As String = GETpName(vpv.vpid)
                    Dim pvname As String = calcProjektKey(pName, vpv.variantName)
                    result.Add(pvname, pvname)

                End If
            Next

        Catch ex As Exception

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
                                                 ByVal storedEarliest As DateTime, ByVal storedLatest As DateTime) As SortedList(Of DateTime, clsProjekt)

        Dim result As New SortedList(Of DateTime, clsProjekt)

        Try

            Dim zwischenResult As New SortedList(Of DateTime, clsProjektWebShort)
            Dim vpid As String = ""
            storedLatest = storedLatest.ToUniversalTime()
            storedEarliest = storedEarliest.ToUniversalTime()

            ' VPID zu Projekt projectName holen vom WebServer/DB
            vpid = GETvpid(projectname)

            If vpid <> "" Then

                ' von der Variante variantName des Projektes vpid alle Versionen holen

                'Dim allVPv As New List(Of clsProjektWebLong)
                'allVPv = GETallVPvLong(vpid, variantName)
                'For Each vpv In allVPv
                '    If storedEarliest < vpv.timestamp And vpv.timestamp < storedLatest Then
                '        'Dim hproj As New clsProjekt
                '        'vpv.copyto(hproj)
                '        'result.Add(vpv.timestamp, hproj)
                '    End If                
                'Next

                Dim allVPv As New List(Of clsProjektWebShort)
                allVPv = GETallVPvShort(vpid, variantName)

                ' einschränken auf alle versionen in dem angegebenen Zeitraum
                For Each vpv In allVPv
                    If storedEarliest < vpv.timestamp And vpv.timestamp < storedLatest Then
                        zwischenResult.Add(vpv.timestamp, vpv)
                    End If
                Next

                ' zu den ausgewählten VPvs nun das Long-Projekt holen
                For Each kvp As KeyValuePair(Of DateTime, clsProjektWebShort) In zwischenResult
                    Dim webProj As List(Of clsProjektWebLong) = GETallVPvLong(kvp.Value.vpid, kvp.Value._id)
                    Dim hproj As New clsProjekt
                    webProj.Item(0).copyto(hproj)
                    result.Add(kvp.Key, hproj)
                Next


            End If
        Catch ex As Exception

        End Try
        retrieveProjectHistoryFromDB = result

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
                                                     ByVal stored As DateTime, ByVal userName As String) As Boolean

        deleteProjectTimestampFromDB = False

    End Function


    ''' <summary>
    ''' holt die erste beauftragte Version des Projects 
    ''' immer mit Variant-Name = ""
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <returns></returns>
    Public Function retrieveFirstContractedPFromDB(ByVal projectname As String) As clsProjekt

        Dim hproj As New clsProjekt

        Try
            Dim vpid As String = ""

            ' VPID zu Projekt projectName holen vom WebServer/DB
            vpid = GETvpid(projectname)

            If vpid <> "" Then

                Dim resultColl As New SortedList(Of DateTime, String)
                Dim allVPv As New List(Of clsProjektWebShort)
                allVPv = GETallVPvShort(vpid)

                For Each vpv As clsProjektWebShort In allVPv
                    If vpv.status = ProjektStatus(PTProjektStati.beauftragt) Then
                        resultColl.Add(vpv.timestamp, vpv._id)
                    End If
                Next

                ' get specific VisboProjectVersion vpvid
                Dim hresult As New List(Of clsProjektWebLong)
                Dim vpvid As String = ""
                If resultColl.count >= 0 Then
                    vpvid = resultColl.ElementAt(0).Value
                End If

                hresult = GETallVPvLong(vpid:="", vpvid:=vpvid)
                If hresult.Count >= 0 Then
                    hresult.Item(0).copyto(hproj)
                Else
                    hproj = Nothing
                End If

            Else
                hproj = Nothing
            End If

        Catch ex As Exception
            hproj = Nothing
        End Try

        retrieveFirstContractedPFromDB = hproj

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
    Public Function checkChgPermission(ByVal pName As String, ByVal vName As String, ByVal userName As String, Optional type As Integer = 0) As Boolean

        'Dim result As Boolean = False
        Dim result As Boolean = True

        Try




            ''clsWriteProtectionItemDB wpItemDB = New clsWriteProtectionItemDB();

            ''    var Filter() = Builders < clsWriteProtectionItemDB > .Filter.Eq("pName", pName) &
            ''                 Builders < clsWriteProtectionItemDB > .Filter.Eq("vName", vName) &
            ''                 Builders < clsWriteProtectionItemDB > .Filter.Eq("type", type);
            ''    //var sort = Builders<clsWriteProtectionItemDB>.Sort.Ascending("pvName");

            ''    bool alreadyExisting = CollectionWriteProtections.AsQueryable < clsWriteProtectionItemDB > ()
            ''                   .Any(wp >= wp.pName == pName && wp.vName == vName && wp.type == type);

            ''    If (alreadyExisting) Then
            ''                        {

            ''        wpItemDB = CollectionWriteProtections.Find(Filter).ToList().Last();
            ''        //var fresult = CollectionWriteProtections.Find(filter).ToList();
            ''        If (wpItemDB.isProtected) Then
            ''                                {
            ''            Return (wpItemDB.userName == userName);   
            ''        }
            ''        Else
            ''        {
            ''            Return True;
            ''        }

            ''    }
            ''    Else
            ''    {
            ''        Return True;
            ''    }
            ''}

            ''Catch (Exception)
            ''{

            ''    Return False;

            ''}
        Catch ex As Exception

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
    Public Function getWriteProtection(ByVal pName As String, ByVal vName As String, Optional type As Integer = 0) As clsWriteProtectionItem
        getWriteProtection = Nothing
    End Function


    ''' <summary>
    ''' setzt für das entsprechende Item das Flag, dass es geschützt ist
    ''' gibt true zurück, wenn die Aktion erfolgreich war, false andernfalls
    ''' </summary>
    ''' <param name="wpItem"></param>
    ''' <returns></returns>
    Public Function setWriteProtection(ByVal wpItem As clsWriteProtectionItem) As Boolean
        Dim result As Boolean = False

        Try
            Dim pname As String = Projekte.getPnameFromKey(wpItem.pvName)
            Dim vname As String = Projekte.getPnameFromKey(wpItem.pvName)
            Dim vpid As String = GETvpid(pname)
            If vpid <> "" Then
                result = POSTVPLock(vpid, vname)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in setWriteProtection: " & ex.Message)
        End Try

        setWriteProtection = result
    End Function



    ''' <summary>
    '''  Alle Portfolios(Constellations) aus der Datenbank holen
    '''  Das Ergebnis dieser Funktion ist eine Liste (String, clsConstellation) 
    ''' </summary>
    ''' <returns></returns>
    Public Function retrieveConstellationsFromDB() As clsConstellations
        retrieveConstellationsFromDB = Nothing
    End Function


    ''' <summary>
    ''' Speichert ein Multiprojekt-Szenario in der Datenbank
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Public Function storeConstellationToDB(ByVal c As clsConstellation) As Boolean
        storeConstellationToDB = False
    End Function

    ''' <summary>
    ''' Löschen des Portfolios  aus der Datenbank
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Public Function removeConstellationFromDB(ByVal c As clsConstellation) As Boolean
        removeConstellationFromDB = False
    End Function



    ''' <summary>
    '''  speichert einen Filter mit Namen 'name' in der Datenbank
    ''' </summary>
    ''' <param name="ptFilter"></param>
    ''' <param name="selfilter"></param>
    ''' <returns></returns>
    Public Function storeFilterToDB(ByVal ptFilter As clsFilter, ByRef selfilter As Boolean) As Boolean
        storeFilterToDB = False
    End Function



    ''' <summary>
    ''' Alle Abhängigkeiten aus der Datenbank lesen
    ''' und als Ergebnis ein Liste von Abhängigkeiten zurückgeben
    ''' </summary>
    ''' <returns></returns>
    Public Function retrieveDependenciesFromDB() As clsDependencies
        retrieveDependenciesFromDB = Nothing
    End Function



    ''' <summary>
    ''' holt von allen Projekt-Varianten in AlleProjekte die Write-Protections
    ''' </summary>
    ''' <param name="AlleProjekte"></param>
    ''' <returns></returns>
    Public Function retrieveWriteProtectionsFromDB(ByVal AlleProjekte As clsProjekteAlle) As SortedList(Of String, clsWriteProtectionItem)
        retrieveWriteProtectionsFromDB = New SortedList(Of String, clsWriteProtectionItem)
    End Function


    ''' <summary>
    ''' löst von allen Projekt-Varianten des Users user die nonpermanent writeProtections
    ''' </summary>
    ''' <param name="user"></param>
    ''' <returns></returns>
    Public Function cancelWriteProtections(ByVal user As String) As Boolean

        Dim result As Boolean = False

        Try
            ' alle vp des aktuellen Users und aktuellen vc holen
            Dim vplist As New List(Of clsVP)
            vplist = GETallVP(aktVCid)

            For Each vp As clsVP In vplist

                ' holt zu der vpid die Varianten aus vpv Collection
                Dim variantToProj As List(Of clsProjektWebShort) = GETallVPvShort(vp._id,, Date.Now)

                ' Lock löschen für jede Variante des Projektes mit vpid
                For Each vTp As clsProjektWebShort In variantToProj
                    result = result And DELETEVPLock(vp._id, vTp.variantName)
                Next
            Next

        Catch ex As Exception
            Throw New ArgumentException("Fehler in cancelWriteProtections:" & ex.Message)
        End Try

        cancelWriteProtections = result
    End Function


    ''' <summary>
    ''' liest alle Filter aus der Datenbank 
    ''' </summary>
    ''' <param name="selfilter"></param>
    ''' <returns></returns>
    Public Function retrieveAllFilterFromDB(ByVal selfilter As Boolean) As SortedList(Of String, clsFilter)
        retrieveAllFilterFromDB = New SortedList(Of String, clsFilter)
    End Function


    ''' <summary>
    ''' löscht einen bestimmten Filter aus der Datenbank
    ''' </summary>
    ''' <param name="filter"></param>
    ''' <returns></returns>
    Public Function removeFilterFromDB(ByVal filter As clsFilter) As Boolean

        removeFilterFromDB = False

    End Function

    ''' <summary>
    ''' liest die Rollendefinitionen aus der Datenbank 
    ''' </summary>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveRolesFromDB(ByVal storedAtOrBefore As DateTime) As clsRollen
        Dim result As New clsRollen()
        Try
            If storedAtOrBefore <= Date.MinValue Then
                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime()
            Else
                storedAtOrBefore = storedAtOrBefore.ToUniversalTime()
            End If

            Dim allRoles As New List(Of clsVCrole)
            ' Alle in der DB-vorhandenen Rollen mit timestamp <= refdate wäre wünschenswert
            allRoles = GETallVCrole(aktVCid)

            For Each role As clsVCrole In allRoles
                Dim roleDef As New clsRollenDefinition
                role.copyTo(roleDef)
                result.Add(roleDef)
            Next

            ' hier werden die topLevelNodeIDs zusammen gesammelt
            result.buildTopNodes()

        Catch ex As Exception

        End Try
        retrieveRolesFromDB = result

    End Function



    ''' <summary>
    '''  liest die Kostenartdefinitionen aus der Datenbank 
    ''' </summary>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveCostsFromDB(ByVal storedAtOrBefore As DateTime) As clsKostenarten

        Dim result As New clsKostenarten()
        Try
            If storedAtOrBefore <= Date.MinValue Then
                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime()
            Else
                storedAtOrBefore = storedAtOrBefore.ToUniversalTime()
            End If

            Dim allCosts As New List(Of clsVCcost)
            ' Alle in der DB-vorhandenen Rollen mit timestamp <= refdate wäre wünschenswert
            allCosts = GETallVCcost(aktVCid)

            For Each cost As clsVCcost In allCosts
                Dim costDef As New clsKostenartDefinition
                cost.copyTo(costDef)
                result.Add(costDef)
            Next


        Catch ex As Exception

        End Try

        retrieveCostsFromDB = result

    End Function


    ''' <summary>
    ''' holt zu dem Projekt projectName die zugehörige vpid vom Server
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <returns></returns>
    Private Function GETvpid(ByVal projectName As String) As String

        Dim vpid As String = ""


        Try
            ' Alle VisboProjects über Server von WebServer/DB holen
            Dim anzLoop As Integer = 0
            'Dim allVP As New List(Of clsVP)
            While (vpid = "" And anzLoop <= 2)

                If VPs.Count > 0 Then
                    ' Id zu angegebenen Projekt herausfinden
                    For Each vp As clsVP In VPs
                        If vp.name = projectName Then
                            vpid = vp._id
                            Exit For
                        End If
                    Next
                End If
                If vpid = "" Then
                    VPs = GETallVP(aktVCid)
                End If

                anzLoop = anzLoop + 1
            End While

        Catch ex As Exception

        End Try

        GETvpid = vpid

    End Function

    ''' <summary>
    ''' holt zu dem Projekt mit der Id vpid den zugehörigen Projektnamen vom Server
    ''' </summary>
    ''' <param name="vpid"></param>
    ''' <returns></returns>
    Private Function GETpName(ByVal vpid As String) As String

        Dim pName As String = ""


        Try
            ' Alle VisboProjects über Server von WebServer/DB holen
            Dim anzLoop As Integer = 0
            'Dim allVP As New List(Of clsVP)
            While (pName = "" And anzLoop <= 2)

                If VPs.Count > 0 Then
                    ' pName zu angegebene vpid herausfinden
                    For Each vp As clsVP In VPs
                        If vp._id = vpid Then
                            pName = vp.name
                            Exit For
                        End If
                    Next
                End If
                If vpid = "" Then
                    VPs = GETallVP(aktVCid)
                End If

                anzLoop = anzLoop + 1
            End While

        Catch ex As Exception

        End Try

        GETpName = pName

    End Function

    ''' <summary>
    ''' holt zu dem VisboCenter vcName die zugehörige vcid vom Server
    ''' </summary>
    ''' <param name="vcName"></param>
    ''' <returns></returns>
    Private Function GETvcid(ByVal vcName As String) As String

        Dim vcid As String = ""


        Try
            ' Alle VisboProjects über Server von WebServer/DB holen
            Dim anzLoop As Integer = 0
            'Dim allVP As New List(Of clsVP)
            While (vcid = "" And anzLoop <= 2)

                If VCs.Count > 0 Then
                    ' Id zu angegebenen Projekt herausfinden
                    For Each vc As clsVC In VCs
                        If vc.name = vcName Then
                            vcid = vc._id
                            Exit For
                        End If
                    Next
                End If
                If vcid = "" Then
                    VCs = GETallVC("")
                End If

                anzLoop = anzLoop + 1
            End While

        Catch ex As Exception

        End Try

        GETvcid = vcid

    End Function


    ''' <summary>
    ''' Holt  VisboCenter mit Name vcName
    ''' </summary>
    ''' <param name="vcName"></param>
    ''' <returns>VisboCenter mit allen Eigenschaften</returns>
    Private Function GETallVC(ByVal vcName As String) As List(Of clsVC)

        Dim result As New List(Of clsVC)

        Try
            Dim serverUriString As String
            Dim typeRequest As String = "/vc"

            ' URL zusammensetzen
            serverUriString = serverUriName & typeRequest
            Dim serverUri As New Uri(serverUriString)

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim Antwort As String
            Dim webVCantwort As clsWebVC
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                webVCantwort = JsonConvert.DeserializeObject(Of clsWebVC)(Antwort)
            End Using

            If webVCantwort.state = "success" Then
                ' Call MsgBox(webVCantwort.message & vbCrLf & "es existieren " & webVCantwort.vc.Count & "VisboCenters")
                result = webVCantwort.vc
            Else
                Call MsgBox(webVCantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in GETallVC: " & ex.Message)
        End Try

        GETallVC = result

    End Function


    ''' <summary>
    ''' Holt alle VisboProject zu dem VisboCenter vcid
    ''' </summary>
    ''' <param name="vcid">vcid = "": es werden alle VisboProjects dieses Users geholt
    '''                    sonst die visboProjects vom Visbocenter vcid</param>
    ''' <returns>Liste der VisboProjects</returns>
    Private Function GETallVP(ByVal vcid As String) As List(Of clsVP)

        Dim result As New List(Of clsVP)

        Try
            Dim serverUriString As String
            Dim typeRequest As String = "/vp"

            ' URL zusammensetzen
            If vcid = "" Then
                serverUriString = serverUriName & typeRequest
            Else
                serverUriString = serverUriName & typeRequest & "?vcid=" & vcid
            End If
            Dim serverUri As New Uri(serverUriString)

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim Antwort As String
            Dim webVPantwort As clsWebVP = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                webVPantwort = JsonConvert.DeserializeObject(Of clsWebVP)(Antwort)
            End Using

            If webVPantwort.state = "success" Then
                ' Call MsgBox(webVPantwort.message & vbCrLf & "aktueller User hat " & webVPantwort.vp.Count & "VisboProjects")

                result = webVPantwort.vp
            Else
                Call MsgBox(webVPantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in PTWebRequest: " & ex.Message)
        End Try

        GETallVP = result

    End Function


    ''' <summary>
    ''' holt zu einer vpid alle VisboProjectsVersionen, wenn ein VarianteName angegeben ist, werden alle Versionen dieser Variante geholt
    ''' bei gegebenen storedAtorBefore nur die neueste Version zu diesem Datum
    ''' </summary>
    ''' <param name="vpid"></param>
    ''' <param name="variantName"></param>
    ''' <param name="storedAtorBefore"></param>
    ''' <returns></returns>
    Private Function GETallVPvShort(ByVal vpid As String,
                              Optional ByVal variantName As String = "",
                              Optional ByVal storedAtorBefore As Date = Nothing) As List(Of clsProjektWebShort)

        Dim result As New List(Of clsProjektWebShort)
        Try

            Dim typeRequest As String = "/vpv"
            Dim serverUriString As String = serverUriName & typeRequest

            If vpid <> "" Then
                serverUriString = serverUriString & "?vpid=" & vpid
            End If

            If storedAtorBefore > Date.MinValue Then
                serverUriString = serverUriString & "&refDate=" & storedAtorBefore.Date.ToString
            End If

            If variantName <> "" Then
                serverUriString = serverUriString & "&variantName=" & variantName
            End If

            Dim serverUri As New Uri(serverUriString)

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim Antwort As String
            Dim webVPvAntwort As clsWebVPv
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                webVPvAntwort = JsonConvert.DeserializeObject(Of clsWebVPv)(Antwort)
            End Using

            If webVPvAntwort.state = "success" Then
                ' Call MsgBox(webVPvAntwort.message & vbCrLf & "aktueller User hat " & webVPvAntwort.vpv.Count & " VisboProjectsVersions")
                result = webVPvAntwort.vpv
            Else
                Throw New ArgumentException(webVPvAntwort.state & ": " & webVPvAntwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in GETallVPvShort: " & ex.Message)
        End Try

        GETallVPvShort = result

    End Function

    ''' <summary>
    ''' holt zu einer vpid alle VisboProjectsVersionen, wenn ein VarianteName angegeben ist, werden alle Versionen dieser Variante geholt
    ''' bei gegebenen storedAtorBefore nur die neueste Version zu diesem Datum
    ''' </summary>
    ''' <param name="vpid"></param>
    ''' <param name="variantName"></param>
    ''' <param name="storedAtorBefore"></param>
    ''' <returns></returns>
    Private Function GETallVPvLong(ByVal vpid As String,
                                   Optional vpvid As String = "",
                                   Optional ByVal variantName As String = "",
                                   Optional ByVal storedAtorBefore As Date = Nothing) As List(Of clsProjektWebLong)

        Dim result As New List(Of clsProjektWebLong)
        Try

            Dim typeRequest As String = "/vpv"
            Dim serverUriString As String = serverUriName & typeRequest

            If vpvid <> "" Then
                serverUriString = serverUriString & "/" & vpvid
            Else
                serverUriString = serverUriName & typeRequest & "?vpid=" & vpid

                If Not IsNothing(storedAtorBefore) Then
                    serverUriString = serverUriString & "&refDate=" & storedAtorBefore.Date.ToString
                End If

                If variantName <> "" Then
                    serverUriString = serverUriString & "&variantName=" & variantName
                End If

                ' es wird die Long-Version einer VisboProjectVersion angefordert
                serverUriString = serverUriString & "&longList"
            End If


            Dim serverUri As New Uri(serverUriString)

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim Antwort As String
            Dim webVPvAntwort As clsWebLongVPv
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                webVPvAntwort = JsonConvert.DeserializeObject(Of clsWebLongVPv)(Antwort)
            End Using

            If webVPvAntwort.state = "success" Then
                ' Call MsgBox(webVPvAntwort.message & vbCrLf & "aktueller User hat " & webVPvAntwort.vpv.Count & " VisboProjectsVersions")
                result = webVPvAntwort.vpv
            Else
                Throw New ArgumentException(webVPvAntwort.state & ": " & webVPvAntwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in GETallVPvLong: " & ex.Message)
        End Try

        GETallVPvLong = result

    End Function

    ''' <summary>
    ''' Holt alle Rollen (vcrole) zu dem VisboCenter vcid
    ''' </summary>
    ''' <param name="vcid">vcid = "": es werden alle Rollen vom Visbocenter vcid  geholt</param>
    '''                    
    ''' <returns>Liste der Rollen</returns>
    Private Function GETallVCrole(ByVal vcid As String) As List(Of clsVCrole)

        Dim result As New List(Of clsVCrole)

        Try
            Dim serverUriString As String
            Dim typeRequest As String = "/vc"

            ' URL zusammensetzen
            If vcid = "" Then
                serverUriString = serverUriName & typeRequest
            Else
                serverUriString = serverUriName & typeRequest & "/" & vcid
            End If
            serverUriString = serverUriString & "/role"

            Dim serverUri As New Uri(serverUriString)

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim Antwort As String
            Dim webVCroleantwort As clsWebVCrole = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                webVCroleantwort = JsonConvert.DeserializeObject(Of clsWebVCrole)(Antwort)
            End Using

            If webVCroleantwort.state = "success" Then
                ' Call MsgBox(webVPantwort.message & vbCrLf & "aktueller User hat " & webVPantwort.vp.Count & "VisboProjects")

                result = webVCroleantwort.vcrole
            Else
                Call MsgBox(webVCroleantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in GETallVCrole: " & ex.Message)
        End Try

        GETallVCrole = result

    End Function

    ''' <summary>
    ''' Holt alle Kostenarten (vccost) zu dem VisboCenter vcid
    ''' </summary>
    ''' <param name="vcid">vcid = "": es werden alle Kostenarten vom Visbocenter vcid geholt</param>
    ''' <returns>Liste der Kostenarten</returns>
    Private Function GETallVCcost(ByVal vcid As String) As List(Of clsVCcost)

        Dim result As New List(Of clsVCcost)

        Try
            Dim serverUriString As String
            Dim typeRequest As String = "/vc"

            ' URL zusammensetzen
            If vcid = "" Then
                serverUriString = serverUriName & typeRequest
            Else
                serverUriString = serverUriName & typeRequest & "/" & vcid
            End If
            serverUriString = serverUriString & "/cost"

            Dim serverUri As New Uri(serverUriString)

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim Antwort As String
            Dim webVCcostantwort As clsWebVCcost = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                webVCcostantwort = JsonConvert.DeserializeObject(Of clsWebVCcost)(Antwort)
            End Using

            If webVCcostantwort.state = "success" Then
                ' Call MsgBox(webVPantwort.message & vbCrLf & "aktueller User hat " & webVPantwort.vp.Count & "VisboProjects")

                result = webVCcostantwort.vccost
            Else
                Call MsgBox(webVCcostantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in GETallVCrole: " & ex.Message)
        End Try

        GETallVCcost = result

    End Function



    ''' <summary>
    ''' ändert ein VisboProject
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird ein VisboProject geändert. user muss die Rechte haben, das checkt der Server</param>
    ''' <returns>Liste der VisboProjects</returns>
    Private Function PUTOneVP(ByVal vpid As String) As List(Of clsVP)

        Dim result As New List(Of clsVP)

        Try
            Dim serverUriString As String = ""
            Dim typeRequest As String = "/vp"

            ' URL zusammensetzen
            If vpid = "" Then
                Call MsgBox("Fehler beim PUTOneVP")
            Else
                serverUriString = serverUriName & typeRequest & "/" & vpid
            End If
            Dim serverUri As New Uri(serverUriString)

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim Antwort As String
            Dim webVPantwort As clsWebVP = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "PUT")
                Antwort = ReadResponseContent(httpresp)
                webVPantwort = JsonConvert.DeserializeObject(Of clsWebVP)(Antwort)
            End Using

            If webVPantwort.state = "success" Then
                ' Call MsgBox(webVPantwort.message & vbCrLf & "aktueller User hat " & webVPantwort.vp.Count & "VisboProjects")

                result = webVPantwort.vp
            Else
                Call MsgBox(webVPantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in PTWebRequest: " & ex.Message)
        End Try

        PUTOneVP = result

    End Function


    ''' <summary>
    ''' Lockt ein Projekt/variante
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird ein VisboProject geändert. user muss die Rechte haben, das checkt der Server</param>
    ''' <returns>Liste der VisboProjects</returns>
    Private Function POSTVPLock(ByVal vpid As String, ByVal variantName As String) As Boolean

        Dim result As Boolean = False

        Try
            ' URL zusammensetzen
            Dim serverUriString As String = ""
            Dim typeRequest As String = "/vp"

            If vpid = "" Then
                Call MsgBox("Fehler beim POSTVPLock")
            Else
                serverUriString = serverUriName & typeRequest & "/" & vpid
            End If
            Dim serverUri As New Uri(serverUriString)

            ' DATA - Block zusammensetzen
            Dim vplock As New clsVPLock
            vplock.variantName = variantName
            vplock.email = aktUser.email
            vplock.expiresAt = DateAdd(DateInterval.Day, 1.0, Date.Now) ' heute + 1 Tag

            Dim data As Byte() = serverInputDataJson(vplock, "")

            ' Request absetzen
            Dim Antwort As String
            Dim webVPLockantwort As clsWebVPLock = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                webVPLockantwort = JsonConvert.DeserializeObject(Of clsWebVPlock)(Antwort)
            End Using

            If webVPLockantwort.state = "success" Then
                ' Call MsgBox(webVPantwort.message & vbCrLf & "aktueller User hat " & webVPantwort.vp.Count & "VisboProjects")

                result = True
            Else
                Call MsgBox(webVPLockantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in POSTVPLock: " & ex.Message)
        End Try

        POSTVPLock = result

    End Function

    ''' <summary>
    ''' löscht den Lock eines Projektes/variante
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird ein der Lock eines VisboProject gelöscht. user muss die Rechte haben, das checkt der Server</param>
    ''' <returns>true: gelöscht
    '''          false: konnte nicht gelöscht werden</returns>
    Private Function DELETEVPLock(ByVal vpid As String, Optional ByVal variantName As String = "") As Boolean

        Dim result As Boolean = False

        Try
            ' URL zusammensetzen
            Dim typeRequest As String = "/vp"
            Dim serverUriString As String = serverUriName & typeRequest

            If vpid = "" Then
                serverUriString = serverUriString & "/lock"
            Else
                serverUriString = serverUriString & "/" & vpid & "/lock"
            End If
            If variantName <> "" Then
                serverUriString = serverUriString & "?variantName=" & variantName
            End If

            Dim serverUri As New Uri(serverUriString)

            ' DATA - Block zusammensetzen
            Dim vplock As New clsVPLock
            vplock.variantName = variantName
            vplock.email = aktUser.email
            vplock.expiresAt = DateAdd(DateInterval.Day, 1.0, Date.Now) ' heute + 1 Tag

            Dim data As Byte() = serverInputDataJson(vplock, "")

            ' Request absetzen
            Dim Antwort As String
            Dim webVPLockantwort As clsWebVPlock = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "DELETE")
                Antwort = ReadResponseContent(httpresp)
                webVPLockantwort = JsonConvert.DeserializeObject(Of clsWebVPlock)(Antwort)
            End Using

            If webVPLockantwort.state = "success" Then
                ' Call MsgBox(webVPantwort.message & vbCrLf & "aktueller User hat " & webVPantwort.vp.Count & "VisboProjects")
                result = True
            Else
                Call MsgBox(webVPLockantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException("Fehler in DELETEVPLock: " & ex.Message)
        End Try

        DELETEVPLock = result

    End Function
End Class

