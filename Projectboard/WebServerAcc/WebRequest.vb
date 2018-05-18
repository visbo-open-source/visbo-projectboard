
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

    Private token As String = ""
    Public webVCs As clsWebVC = Nothing
    Public aktVC As clsWebVC = Nothing
    Public webVPs As clsWebVP = Nothing
    Public VPs As List(Of clsVP) = Nothing
    Public aktVP As clsWebVP = Nothing
    Public webVPvs As clsWebVPv = Nothing
    Public aktVPv As clsWebLongVPv = Nothing



    ''' <summary>
    ''' Sendet einen Request vom Typ method an den Server. Außerdem wird hier auch die Antwort empfangen und an die aufrufenden Routine zurückgegeben
    ''' </summary>
    ''' <param name="uri">Url fur den REst-Request</param>
    ''' <param name="data">Daten für die Aufrufe von POST/PUT</param>
    ''' <param name="method">Typ des Rest-Request  GET/POST/PUT/DELETE</param>
    Public Function GetRestServerResponse(ByVal uri As Uri, ByVal data As Byte(), ByVal method As String) As HttpWebResponse
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
                    Call MsgBox("Fehler bei GetRequestStream:  " & ex.Message)
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

    Public Function ReadResponseContent(ByRef resp As HttpWebResponse) As String
        'Private Function ReadResponseContent(ByRef resp As HttpWebResponse) As String

        If IsNothing(resp) Then
            Throw New ArgumentNullException("resp")
        Else
            Using sr As New StreamReader(resp.GetResponseStream)
                Return sr.ReadToEnd()
            End Using
        End If

    End Function


    ''' <summary>
    ''' diese Funktion konvertiert die Struktur, die für diesen Server-Request benötigt wird (type) in ein ByteArray im Json-Format
    ''' </summary>
    ''' <param name="dataClass"></param>
    ''' <param name="type"></param>
    ''' <returns>Object</returns>
    Public Function serverInputDataJson(ByVal dataClass As Object, ByVal type As String) As Byte()
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
            Else
                token = ""
                serverUriName = ServerURL
            End If


        Catch ex As Exception
            Call MsgBox("Fehler in PTWebRequestLogin" & typeRequest & ": " & ex.Message)
        End Try

        login = loginOK

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
                For Each shortproj As clsProjektWebShort In allVPv
                    ergebnisCollection.Add(shortproj.timestamp, shortproj.timestamp)
                Next
            End If

        Catch ex As Exception

        End Try

        retrieveZeitstempelFromDB = ergebnisCollection

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
                Dim resultColl As New Collection
                Dim allVPv As New List(Of clsProjektWebShort)
                allVPv = GETallVPvShort(vpid)
                For Each vpv As clsProjektWebShort In allVPv
                    If vpv.status = ProjektStatus(PTProjektStati.beauftragt) Then
                        resultColl.Add(vpv.timestamp, vpv._id)
                    End If
                Next
                Dim hresult As New List(Of clsProjektWebLong)
                Dim anz As Integer = resultColl.Count
                hresult = GETallVPvLong("", resultColl.Item(resultColl.Count))
                hresult.Item(0).copyto(hproj)
            Else
                hproj = Nothing
            End If

        Catch ex As Exception
            hproj = Nothing
        End Try

        retrieveFirstContractedPFromDB = hproj

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
                    If vpid = "" Then
                        VPs = GETallVP("")
                    End If
                End If
                anzLoop = anzLoop + 1
            End While

        Catch ex As Exception

        End Try

        GETvpid = vpid

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
            Dim webVPantwort As clsWebVP
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")

                Antwort = ReadResponseContent(httpresp)
                webVPantwort = JsonConvert.DeserializeObject(Of clsWebVP)(Antwort)
            End Using

            If webVPantwort.state = "success" Then
                Call MsgBox(webVPantwort.message & vbCrLf & "aktueller User hat " & webVPantwort.vp.Count & "VisboProjects")
                ' hier erfolgen nun die weiteren Aktionen mit den angeforderten Daten

                'webVPs = webVPantwort
                result = webVPantwort.vp
            Else
                Call MsgBox(webVPantwort.message)
            End If

        Catch ex As Exception
            Call MsgBox("Fehler in PTWebRequest: " & ex.Message)
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
            Dim serverUriString As String = serverUriName & typeRequest & "?vpid=" & vpid

            If Not IsNothing(storedAtorBefore) Then
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

                Call MsgBox(webVPvAntwort.message & vbCrLf & "aktueller User hat " & webVPvAntwort.vpv.Count & " VisboProjectsVersions")
                ' hier erfolgen nun die weiteren Aktionen mit den angeforderten Daten

                'webVPvs = webVPvAntwort
                result = webVPvAntwort.vpv
            Else
                Call MsgBox(webVPvs.message)
            End If

        Catch ex As Exception
            Call MsgBox("Fehler in PTWebRequest: " & ex.Message)
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
            End If


            ' es wird die Long-Version einer VisboProjectVersion angefordert
            serverUriString = serverUriString & "&longList"

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

                Call MsgBox(webVPvAntwort.message & vbCrLf & "aktueller User hat " & webVPvAntwort.vpv.Count & " VisboProjectsVersions")
                ' hier erfolgen nun die weiteren Aktionen mit den angeforderten Daten

                'webVPvs = webVPvAntwort
                result = webVPvAntwort.vpv
            Else
                Call MsgBox(webVPvs.message)
            End If

        Catch ex As Exception
            Call MsgBox("Fehler in PTWebRequest: " & ex.Message)
        End Try

        GETallVPvLong = result

    End Function
End Class

