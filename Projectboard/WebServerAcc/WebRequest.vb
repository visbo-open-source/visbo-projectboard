
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
    'Public serverUriName As String = "http://localhost:3484"

    Private serverUriName As String = ""

    Private aktVCid As String = ""

    Private token As String = ""
    Private VCs As New List(Of clsVC)

    Public VRScache As New clsCache
    ' hierin werden  alle Visbo-Projects und 
    ' die vom Server bereits angeforderten VisboProjectsVersionsgecacht
    '
    ' Private VPs As New SortedList(Of String, clsVP)
    '                                     vpid                  vname    timestamp-Liste, projectshort
    ' Private VPvCache As New SortedList(Of String, SortedList(Of String, clstest))
    ' Private VPvCache As New clsCache


    Private aktUser As clsUserReg = Nothing

    'Private webVCs As clsWebVC = Nothing

    'Private aktVC As clsWebVC = Nothing
    'Private webVPs As clsWebVP = Nothing

    'Private aktVP As clsWebVP = Nothing
    'Private webVPvs As clsWebVPv = Nothing
    'Private aktVPv As clsWebLongVPv = Nothing




    ''' <summary>
    '''  'Verbindung mit der Datenbank aufbauen (mit Angabe von Username und Passwort)
    ''' </summary>
    ''' <param name="ServerURL"></param>
    ''' <param name="databaseName">wird beim Login am Visbo-Rest-Server nicht benötigt</param>
    ''' <param name="username"></param>
    ''' <param name="dbPasswort"></param>
    Public Function login(ByVal ServerURL As String, ByVal databaseName As String, ByVal username As String, ByVal dbPasswort As String) As Boolean

        Dim typeRequest As String = "/token/user/login"
        'Dim typeRequest As String = "/token/user/signup"
        Dim serverUri As New Uri(ServerURL & typeRequest)
        Dim loginOK As Boolean = False
        Dim httpresp_sav As HttpWebResponse

        Try
            Dim user As New clsUserLoginSignup
            user.email = LCase(username)
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
                httpresp_sav = httpresp     ' sichern der Server-Antwort
                loginAntwort = JsonConvert.DeserializeObject(Of clsWebTokenUserLoginSignup)(Antwort)
            End Using

            If awinSettings.visboDebug Then
                Call MsgBox(loginAntwort.message)
            End If

            loginOK = (loginAntwort.state = "success")

            If loginOK Then
                token = loginAntwort.token
                serverUriName = ServerURL
                aktUser = loginAntwort.user
                ' VisboCenterID mit Name = databaseName wird gespeichert
                aktVCid = GETvcid(databaseName)

                If aktVCid = "" Then
                    loginOK = False
                    token = ""
                    If awinSettings.englishLanguage Then
                        Call MsgBox("User don't have access to this VisboCenter!" & vbLf & "Please contact your administrator")
                    Else
                        Call MsgBox("User hat keinen Zugriff zu diesem VisboCenter!" & vbLf & " Bitte kontaktieren Sie ihren Administrator")
                    End If

                End If

            Else
                token = ""
                serverUriName = ServerURL
                aktUser = Nothing
                If awinSettings.visboDebug Then
                    Call MsgBox("( " & CType(httpresp_sav.StatusCode, Integer).ToString & ") : " & httpresp_sav.StatusDescription & " : " & loginAntwort.message)
                End If

            End If


        Catch ex As Exception
            Throw New ArgumentException("Fehler in PTWebRequestLogin" & typeRequest & ": " & ex.Message)
        End Try

        login = loginOK

    End Function

    ''' <summary>
    ''' prüft die Verfügbarkeit der MongoDB bzw. ob ein Login bereits erfolgte, d.h. gültiger token vorhanden
    ''' </summary>
    ''' <returns></returns>
    Public Function pingMongoDb() As Boolean

        Dim result As Boolean = False
        Try
            If token <> "" Then
                Dim vcid As String = GETvcid(awinSettings.databaseName)
                result = (vcid <> "")
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        pingMongoDb = result
    End Function

    ''' <summary>
    ''' über Email setzen einen neuen Passwortes; geht nur beim Server
    ''' </summary>
    ''' <returns></returns>
    Public Function pwforgotten(ByVal ServerURL As String, ByVal databaseName As String, ByVal username As String) As Boolean

        Dim result As Boolean = False
        Try
            result = POSTpwforgotten(ServerURL, databaseName, username)

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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
    Public Function projectNameAlreadyExists(ByVal projectname As String, ByVal variantname As String, ByVal storedAtorBefore As DateTime) As Boolean

        Dim result As Boolean = False

        Try
            If storedAtorBefore <= Date.MinValue Then
                storedAtorBefore = DateTime.Now.AddDays(1).ToUniversalTime()
            Else
                storedAtorBefore = storedAtorBefore.ToUniversalTime()
            End If

            Dim vpid As String = ""

            vpid = GETvpid(projectname)._id

            If vpid <> "" And variantname <> "" Then
                ' nachsehen, ob im VisboProject diese Variante zum Zeitpunkt storedAtorBefore bereits created war
                For Each vpVar As clsVPvariant In VRScache.VPsN(projectname).Variant
                    If vpVar.variantName = variantname Then
                        If vpVar.createdAt <= storedAtorBefore Then
                            result = True
                            Exit For
                        End If
                    End If
                Next
            Else
                result = (vpid <> "")
            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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

            ' alle VisboProjectVersions vom Server anfordern
            ' ur:08.06.2018: wird in globale Variable gecacht: Dim allVPv As New List(Of clsProjektWebShort)

            Dim allVPv As New List(Of clsProjektWebShort)
            allVPv = GETallVPvShort("")

            ' alle vorhandenen Timestamps in der resultCollection sammeln
            Dim sl As New SortedList(Of Date, Date)
            For Each shortproj As clsProjektWebShort In allVPv
                If Not sl.ContainsKey(shortproj.timestamp) Then
                    sl.Add(shortproj.timestamp, shortproj.timestamp)
                End If
            Next

            For i As Integer = sl.Count - 1 To 0 Step -1
                Dim kvp As KeyValuePair(Of DateTime, DateTime) = sl.ElementAt(i)
                resultCollection.Add(kvp.Value.ToUniversalTime())
            Next i

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
    Public Function retrieveZeitstempelFromDB(ByVal pvName As String) As Collection

        Dim ergebnisCollection As New Collection
        'token = ""
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
            vpid = GETvpid(projectName)._id

            If vpid <> "" Then
                ' gewünschte Variante vom Server anfordern
                Dim allVPv As New List(Of clsProjektWebShort)
                allVPv = GETallVPvShort(vpid, , variantName)

                ' alle vorhandenen Timestamps zu einem pvName in die ErgebnisCollection sammeln
                Dim sl As New SortedList(Of Date, Date)
                For Each shortproj As clsProjektWebShort In allVPv
                    If Not sl.ContainsKey(shortproj.timestamp) Then
                        sl.Add(shortproj.timestamp, shortproj.timestamp)
                    End If
                Next

                For i As Integer = sl.Count - 1 To 0 Step -1
                    Dim kvp As KeyValuePair(Of DateTime, DateTime) = sl.ElementAt(i)
                    '???: ergebnisCollection.Add(kvp.Value.ToUniversalTime)
                    ergebnisCollection.Add(kvp.Value.ToLocalTime)
                Next i

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveZeitstempelFromDB = ergebnisCollection

    End Function
    ''' <summary>
    ''' bringt für die angegebene Projekt-Variante alle Zeitstempel in absteigender Sortierung zurück 
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <returns>Collection, absteigend sortiert</returns>
    Public Function retrieveZeitstempelFirstLastFromDB(ByVal pvName As String) As Collection

        Dim ergebnisCollection As New Collection
        'token = ""
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
            vpid = GETvpid(projectName)._id

            If vpid <> "" Then

                Dim hresultFirst As New List(Of clsProjektWebShort)

                hresultFirst = GETallVPvShort(vpid, vpvid:="", status:="", refNext:=True, variantName:=variantName, storedAtorBefore:=Nothing)

                Dim anzResult As Integer = hresultFirst.Count
                If anzResult >= 0 Then
                    ergebnisCollection.Add(hresultFirst.Item(anzResult - 1).timestamp)
                End If

                Dim hresultLast As New List(Of clsProjektWebShort)

                hresultLast = GETallVPvShort(vpid, , , refNext:=False, variantName:=variantName, storedAtorBefore:=Date.Now.ToUniversalTime)

                If hresultLast.Count >= 0 Then
                    ergebnisCollection.Add(hresultLast.Item(0).timestamp)
                End If

            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveZeitstempelFirstLastFromDB = ergebnisCollection

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
        Dim diffRCBeginn As Date = Date.Now
        Dim diffRC As Long
        Dim diffCopy As Long
        'Dim diff3 As Long
        Try
            Dim hproj As New clsProjekt

            ' da in der Datenbank alle DateTime im UTC gespeichert sind, muss hier auch dieses Format verwendet werden
            Dim aktDate As Date = Date.Now.ToUniversalTime()



            storedLatest = storedLatest.ToUniversalTime()
            storedEarliest = storedEarliest.ToUniversalTime()
            zeitraumStart = zeitraumStart.ToUniversalTime()
            zeitraumEnde = zeitraumEnde.ToUniversalTime()


            ' Kein Projekt  angegeben. es werden alle Projekte im angebenen Zeitraum zurückgegeben

            If projectname = "" Then


                VRScache.VPsN = GETallVP(aktVCid, ptPRPFType.project)

                Dim VisboPv_all As New List(Of clsProjektWebLong)
                VisboPv_all = GETallVPvLong("", , , , variantName, aktDate)

                diffRC = DateDiff(DateInterval.Second, diffRCBeginn, Date.Now)

                Dim copyBeginn As Date = Date.Now

                For Each webProj As clsProjektWebLong In VisboPv_all

                    If (webProj.startDate <= zeitraumEnde And
                                webProj.endDate >= zeitraumStart And
                                webProj.timestamp <= storedLatest) Then

                        hproj = New clsProjekt

                        webProj.copyto(hproj)
                        Dim a As Integer = hproj.dauerInDays
                        Dim key As String = Projekte.calcProjektKey(hproj)
                        If Not result.ContainsKey(key) Then
                            result.Add(key, hproj)
                        End If

                    End If

                Next

                diffCopy = DateDiff(DateInterval.Second, copyBeginn, Date.Now)

                '' ur: 2018.11.14: das Holen aller Projekte und Varianten einzeln verursacht zu lange Antwortzeit
                ''
                '' schleife über alle VisboProjects
                'For Each kvp As KeyValuePair(Of String, clsVP) In VRScache.VPsN

                '    Dim vpid As String = kvp.Value._id

                '    If vpid <> "" Then
                '        ' gewünschten Varianten vom Server anfordern
                '        Dim allVPv As New List(Of clsProjektWebLong)
                '        allVPv = GETallVPvLong(vpid, , variantName, aktDate)

                '        For Each webProj As clsProjektWebLong In allVPv

                '            If (webProj.startDate <= zeitraumEnde And
                '                webProj.endDate >= zeitraumStart And
                '                webProj.timestamp <= storedLatest) Then

                '                hproj = New clsProjekt

                '                webProj.copyto(hproj)
                '                Dim a As Integer = hproj.dauerInDays
                '                Dim key As String = Projekte.calcProjektKey(hproj)
                '                If Not result.ContainsKey(key) Then
                '                    result.Add(key, hproj)
                '                End If

                '            End If

                '        Next
                '    Else
                '        ' kann eigentlich nicht vorkommen
                '    End If

                'Next

            Else
                '  Projekt angegeben: d.h. es werden alle Timestamps der übergebenen Projekt-Variante zurückgegeben
                Dim vpid As String = GETvpid(projectname)._id
                If vpid <> "" Then
                    ' gewünschten Varianten vom Server anfordern
                    Dim allVPv As New List(Of clsProjektWebLong)
                    allVPv = GETallVPvLong(vpid, , , , variantName, storedLatest)

                    For Each webProj As clsProjektWebLong In allVPv
                        If webProj.timestamp >= storedEarliest Then

                            hproj = New clsProjekt

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
            Throw New ArgumentException(ex.Message)
        End Try



        retrieveProjectsFromDB = result

        Call MsgBox("RestTime: " & diffRC.ToString & vbLf & "CopyTime: " & diffCopy.ToString)

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

        storedAtOrBefore = storedAtOrBefore.ToUniversalTime

        Try
            Dim hproj As New clsProjekt
            Dim vpid As String = ""
            vpid = GETvpid(projectname)._id

            If vpid <> "" Then
                ' gewünschte Variante vom Server anfordern
                Dim allVPv As New List(Of clsProjektWebLong)
                allVPv = GETallVPvLong(vpid, , , , variantname, storedAtOrBefore)
                If allVPv.Count > 0 Then
                    Dim webProj As clsProjektWebLong = allVPv.ElementAt(0)
                    webProj.copyto(hproj)
                    result = hproj
                End If

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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

                    Dim vp As New clsVP

                    Dim vpid As String = GETvpid(oldName)._id
                    If VRScache.VPsN.ContainsKey(oldName) Then

                        vp = VRScache.VPsN(oldName)
                        If VRScache.VPsN.Remove(oldName) Then
                            vp._id = vpid
                            vp.name = newName
                            VRScache.VPsN.Add(newName, vp)
                        End If

                    End If

                    Dim vpList As List(Of clsVP) = PUTOneVP(vpid, vp)
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
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try

            Dim webVP As New clsWebVP
            Dim vpErg As New List(Of clsVP)
            Dim data() As Byte

            Dim pname As String = projekt.name
            Dim vname As String = projekt.variantName

            Dim aktvp As clsVP = GETvpid(pname, projekt.projectType)
            Dim vpid As String = aktvp._id
            Dim storedVP As Boolean = (vpid <> "")

            If Not storedVP Then

                Dim typeRequest As String = "/vp"
                Dim serverUriString As String = serverUriName & typeRequest
                Dim serverUri As New Uri(serverUriString)

                Dim VP As New clsVP
                'ur: 8.11.2018: testweise: nach Telefonat mit ms soll der User, der hier zugriff hat vom ReST-Server eingetragen werden.
                'Dim user As New clsUser
                'user.email = aktUser.email
                'user.role = "Admin"
                'VP.users.Add(user)
                VP.name = pname
                VP.vcid = aktVCid
                VP.vpPublic = True
                VP.vpType = projekt.projectType

                vpErg = POSTOneVP(VP)


                If vpErg.Count > 0 Then

                    ' vpErg.ElementAt(0) ist nun das aktuelle VP
                    vpid = vpErg.ElementAt(0)._id
                    aktvp = vpErg.ElementAt(0)
                    storedVP = (vpid <> "")

                Else
                    Throw New ArgumentException("Das VisboProject existiert nicht und konnte auch nicht erzeugt werden!")
                End If

            End If

            ' überprüfen, ob die gewünschte Variante im VisboProject enthalten ist
            Dim storedVPVariant As Boolean = False
            If vname <> "" And aktvp.Variant.Count > 0 Then
                For Each var As clsVPvariant In aktvp.Variant
                    If var.variantName = vname Then
                        storedVPVariant = True
                    End If
                Next
            End If

            ' wenn Variante noch nicht vorhanden, so muss sie angelegt werden
            If Not storedVPVariant Then
                If vname <> "" Then
                    storedVPVariant = POSTVPVariant(vpid, vname)
                Else
                    ' zu diesem Projekt gibt es nur die Standardvariante = > nichts tun
                End If
            End If

            ' Projekt ist bereits in VisboProjects Collection gespeichert, es existiert eine vpid
            If storedVP Then

                ' jetzt muss noch VisboProjectVersion gespeichert werden
                Dim typeRequest As String = "/vpv"
                Dim serverUriString As String = serverUriName & typeRequest
                Dim serverUri As New Uri(serverUriString)


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
                        errcode = CType(httpresp.StatusCode, Integer)
                        errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                        storeAntwort = JsonConvert.DeserializeObject(Of clsWebLongVPv)(Antwort)
                    End Using

                    If errcode = 200 Then

                        result = (storeAntwort.state = "success")
                        result = True

                    Else

                        ' Fehlerbehandlung je nach errcode
                        Dim statError As Boolean = errorHandling_withBreak("POSTOneVPv", errcode, errmsg & " : " & storeAntwort.message)

                    End If

                End If
            End If

            ' Cache aktualisieren
            VRScache.VPsN = GETallVP(aktVCid, projekt.projectType)

        Catch ex As Exception
            Throw New ArgumentException(ex.Message & ": storeProjectToDB")
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
            vpid = GETvpid(projectName)._id

            ' nun ist sicher die VPs aufgebaut
            Dim vp As clsVP = VRScache.VPsN(projectName)

            ' alle Variantenamen in der Collection sammeln
            For Each vpVar As clsVPvariant In vp.Variant
                ergebnisCollection.Add(vpVar.variantName, vpVar.variantName)
            Next


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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
            zeitraumStart = zeitraumStart.ToUniversalTime()
            zeitraumEnde = zeitraumEnde.ToUniversalTime()

            ''Dim intermediate As New SortedList(Of String, clsVP)

            ''intermediate = GETallVP(aktVCid, ptPRPFType.project)

            ''For Each kvp As KeyValuePair(Of String, clsVP) In intermediate

            ''    Dim vpid As String = kvp.Value._id
            ''    If kvp.Value.Variant.Count >= 0 Then

            ' holt alle Projekte/Variante/versionen mit ReferenzDatum storedatOrBefore
            Dim vpvListe As New List(Of clsProjektWebShort)
            vpvListe = GETallVPvShort("", "", "", False, Nothing, storedAtOrBefore)

            For Each vpv As clsProjektWebShort In vpvListe
                Dim vpType As Integer = GETvpType(vpv.vpid)
                If vpv.startDate <= zeitraumEnde And
                    vpv.endDate >= zeitraumStart And
                    vpType = ptPRPFType.project Then

                    Dim pName As String = GETpName(vpv.vpid)
                    Dim pvname As String = calcProjektKey(pName, vpv.variantName)
                    If Not result.ContainsKey(pvname) Then
                        result.Add(pvname, pvname)
                    End If

                End If
            Next

            'End If

            '    If kvp.Value.Variant.Count > 0 Then

            '        For Each var As clsVPvariant In kvp.Value.Variant

            '            ' holt alle Projekte/Variante/versionen mit ReferenzDatum storedatOrBefore
            '            Dim vpvListe As New List(Of clsProjektWebShort)
            '            vpvListe = GETallVPvShort(vpid, var.variantName, storedAtOrBefore)

            '            For Each vpv As clsProjektWebShort In vpvListe

            '                If vpv.startDate <= zeitraumEnde And
            '                   vpv.endDate >= zeitraumStart Then

            '                    Dim pName As String = GETpName(vpv.vpid)
            '                    Dim pvname As String = calcProjektKey(pName, vpv.variantName)
            '                    If Not result.ContainsKey(pvname) Then
            '                        result.Add(pvname, pvname)
            '                    End If

            '                End If
            '            Next
            '        Next
            '    End If

            'Next 'von intermediate-Schleife

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveProjectVariantNamesFromDB = result

    End Function

    ''' <summary>
    ''' holt Projekt-Namen über Angabe der Projekt-Nummer beim Kunden; 
    ''' kann Null, ein oder mehrere Ergebnis-Einträge enthalten; Liste kommt sortiert nach Projekt-Namen zurück
    ''' </summary>
    ''' <param name="pNRatKD"></param>
    ''' <returns></returns>
    Public Function retrieveProjectNamesByPNRFromDB(ByVal pNRatKD As String) As Collection

        Dim result As New Collection
        Dim interimResult As New Collection

        Try

            Dim vpid As String = ""
            Dim anzLoop As Integer = 0
            'Dim allVP As New List(Of clsVP)
            While (result.Count <= 0 And anzLoop <= 2)

                ' zuerst nur im Cache nachsehen
                For Each kvp As KeyValuePair(Of String, clsVP) In VRScache.VPsId

                    If (kvp.Value.kundennummer = pNRatKD) And
                        (kvp.Value.vpType = ptPRPFType.project) Then

                        ' Projektnamen einsammeln
                        result.Add(kvp.Value.name)
                        ' vpid-eingesammelt, aktuell nicht weiter verwertet
                        interimResult.Add(kvp.Value._id)
                    End If

                Next
                ' im Cache nicht gefunden, also nochmals alle VP des aktVCid holen und nachsehen.
                If result.Count <= 0 Then

                    VRScache.VPsId = GETallVP(aktVCid, ptPRPFType.all)

                End If

                anzLoop = anzLoop + 1

            End While

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveProjectNamesByPNRFromDB = result

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
        storedLatest = storedLatest.ToUniversalTime()
        storedEarliest = storedEarliest.ToUniversalTime()

        Try

            'Dim zwischenResult As New SortedList(Of DateTime, clsProjektWebLong)
            Dim vpid As String = ""


            ' VPID zu Projekt projectName holen vom WebServer/DB
            vpid = GETvpid(projectname)._id

            If vpid <> "" Then

                Dim allVPv As New List(Of clsProjektWebLong)
                allVPv = GETallVPvLong(vpid, , , , variantName)

                ' einschränken auf alle versionen in dem angegebenen Zeitraum
                For Each vpv In allVPv
                    If storedEarliest <= vpv.timestamp And vpv.timestamp <= storedLatest And vpv.variantName = variantName Then
                        'zwischenResult.Add(vpv.timestamp, vpv)
                        Dim hproj As New clsProjekt
                        vpv.copyto(hproj)
                        result.Add(hproj.timeStamp, hproj)
                    End If
                Next

            End If
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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

        Dim result As Boolean = False

        If aktUser.email = userName Then

            stored = stored.ToUniversalTime
            Try
                Dim vpid As String = ""

                ' VPID zu Projekt projectName holen vom WebServer/DB
                vpid = GETvpid(projectname)._id

                If vpid <> "" Then
                    ' gewünschte Variante vom Server anfordern
                    Dim allVPv As New List(Of clsProjektWebShort)
                    allVPv = GETallVPvShort(vpid, , "", False, variantName, stored)
                    If allVPv.Count >= 0 Then
                        If allVPv.Count = 1 Then
                            result = DELETEOneVPv(allVPv.Item(0)._id)
                        Else
                            For Each vpv As clsProjektWebShort In allVPv
                                If vpv.variantName = variantName Then
                                    result = result And DELETEOneVPv(vpv._id)
                                End If
                            Next
                        End If

                    End If

                End If
            Catch ex As Exception
                Throw New ArgumentException(ex.Message)
            End Try
        Else

            Call MsgBox("Fehler in deletProjectTimestampFromDB: User '" & userName & "' darf nicht löschen")

        End If

        deleteProjectTimestampFromDB = result

    End Function


    ''' <summary>
    ''' holt die erste beauftragte Version des Projects 
    ''' immer mit Variant-Name = variantname
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <param name="variantname"></param>
    ''' <returns></returns>
    Public Function retrieveFirstContractedPFromDB(ByVal projectname As String, ByVal variantname As String) As clsProjekt

        Dim hproj As New clsProjekt

        Try
            Dim vpid As String = ""
            Dim vp As clsVP = GETvpid(projectname)

            If Not IsNothing(vp) Then

                ' VPID zu Projekt projectName holen vom WebServer/DB
                vpid = vp._id

                If vpid <> "" Then

                    Dim hresult As New List(Of clsProjektWebLong)

                    hresult = GETallVPvLong(vpid:=vpid, vpvid:="",
                                                status:="beauftragt",
                                                refNext:=True,
                                                variantName:=variantname,
                                                storedAtorBefore:=Nothing)

                    If hresult.Count >= 0 Then
                        hresult.Item(0).copyto(hproj)
                    Else
                        hproj = Nothing
                    End If

                Else
                    hproj = Nothing
                End If
            Else
                hproj = Nothing
            End If


        Catch ex As Exception
            hproj = Nothing
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveFirstContractedPFromDB = hproj

    End Function
    ''' <summary>
    ''' gibt den zum Zeitpunkt zuletzt beauftragten Stand zurück; bei Projekten muss variantNAme = "" sein, bei Summary Projekten VPortfolioName
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <param name="variantname"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveLastContractedPFromDB(ByVal projectname As String,
                                                  ByVal variantname As String,
                                                  ByVal storedAtOrBefore As DateTime) As clsProjekt

        Dim hproj As New clsProjekt

        Try
            If (storedAtOrBefore = Date.MinValue) Then
                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime()
            Else
                storedAtOrBefore = storedAtOrBefore.ToUniversalTime()
            End If

            Dim vpid As String = ""
            Dim vp As clsVP = GETvpid(projectname)

            If Not IsNothing(vp) Then

                ' VPID zu Projekt projectName holen vom WebServer/DB
                vpid = vp._id

                If vpid <> "" Then

                    ' get specific VisboProjectVersion vpvid
                    Dim hresult As New List(Of clsProjektWebLong)


                    hresult = GETallVPvLong(vpid:=vpid, vpvid:="",
                                            status:="beauftragt",
                                            refNext:=False,
                                            variantName:=variantname,
                                            storedAtorBefore:=storedAtOrBefore)

                    If hresult.Count >= 0 Then
                        hresult.Item(0).copyto(hproj)
                    Else
                        hproj = Nothing
                    End If

                Else
                    hproj = Nothing
                End If

            Else
                hproj = Nothing
            End If



        Catch ex As Exception
            hproj = Nothing
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveLastContractedPFromDB = hproj

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
    Public Function checkChgPermission(ByVal pName As String, ByVal vName As String, ByVal userName As String, Optional type As Integer = ptPRPFType.project) As Boolean

        Dim result As Boolean = False

        Try
            ' angepasst: 20180914: ur: type muss im ReST-Server auf unsere Enumeration geändert werden: 
            '           ptPRPFType.portfolio = 1
            '           ptPRPFType.project = 0
            '           ptPRPFType.projectTemplate = 2

            type = ptPRPFType.project
            Dim wpItem As clsWriteProtectionItem = getWriteProtection(pName, vName, type)

            If wpItem.isProtected Then
                result = (wpItem.userName = aktUser.email)
            Else
                result = True
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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
    Public Function getWriteProtection(ByVal pName As String, ByVal vName As String, Optional type As Integer = ptPRPFType.project) As clsWriteProtectionItem
        Dim result As New clsWriteProtectionItem
        Try
            Dim vp As clsVP = GETvpid(pName, type)
            result.pvName = calcProjektKey(pName, vName)
            result.isProtected = False
            result.isSessionOnly = True
            result.permanent = False
            result.lastDateSet = Nothing
            result.lastDateReleased = Nothing
            result.userName = ""
            result.type = type

            If vp.lock.Count > 0 Then
                For Each vplock As clsVPLock In vp.lock

                    If vplock.variantName = vName Then
                        If vplock.expiresAt.ToLocalTime > Date.Now Then
                            result.isProtected = True
                        Else
                            result.isProtected = False
                        End If
                        result.isSessionOnly = True
                        result.lastDateSet = vplock.createdAt.ToLocalTime
                        result.userName = vplock.email
                        Exit For

                    End If
                Next
            End If



        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        getWriteProtection = result

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
            Dim vname As String = Projekte.getVariantnameFromKey(wpItem.pvName)

            Dim aktvp As clsVP = GETvpid(pname)
            Dim vpid As String = aktvp._id
            Dim variantExists As Boolean = False

            For Each var As clsVPvariant In aktvp.Variant
                If var.variantName = vname Then
                    variantExists = True
                    Exit For
                End If
            Next
            If (vpid <> "" And variantExists) Or (vpid <> "" And vname = "") Then

                If wpItem.isProtected Then
                    result = POSTVPLock(vpid, vname)
                Else
                    result = DELETEVPLock(vpid, vname)
                End If

            Else

                result = False

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        setWriteProtection = result
    End Function



    ''' <summary>
    '''  Alle Portfolios(Constellations) aus der Datenbank holen
    '''  Das Ergebnis dieser Funktion ist eine Liste (String, clsConstellation) 
    ''' </summary>
    ''' <returns></returns>
    Public Function retrieveConstellationsFromDB() As clsConstellations

        Dim result As New clsConstellations
        Try

            Dim intermediate As New SortedList(Of String, clsVP)
            Dim timestamp As Date = Date.Now
            Dim c As New clsConstellation

            intermediate = GETallVP(aktVCid, ptPRPFType.portfolio)
            For Each kvp As KeyValuePair(Of String, clsVP) In intermediate


                If kvp.Value.vpType = ptPRPFType.portfolio Then


                    Dim vpid As String = kvp.Value._id
                    Dim portfolioVersions As SortedList(Of Date, clsVPf) = GETallVPf(vpid, timestamp)
                    If portfolioVersions.Count > 0 Then

                        Dim aktPortfolio As clsVPf = portfolioVersions.Last.Value

                        c = clsVPf2clsConstellation(aktPortfolio)

                        If Not IsNothing(c) Then

                            If Not result.Contains(c.constellationName) Then
                                result.Add(c)
                            End If

                        End If

                    End If

                Else
                    ' es gibt keine Portfolios

                End If

            Next

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveConstellationsFromDB = result
    End Function


    ''' <summary>
    ''' Speichert ein Multiprojekt-Szenario in der Datenbank
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Public Function storeConstellationToDB(ByVal c As clsConstellation) As Boolean

        Dim result As Boolean = False

        Try
            Dim vpType As Integer = ptPRPFType.portfolio
            Dim cVPf As New clsVPf
            Dim cVP As New clsVP
            Dim newVP As New List(Of clsVP)
            Dim newVPf As New List(Of clsVPf)

            ' angepasst: 20180914: korrigieren, wenn ReST-Server geändert wurde
            '                       cVP = GETvpid(c.constellationName, vpType:=2)
            cVP = GETvpid(c.constellationName, ptPRPFType.portfolio)


            'cVPf = clsConst2clsVPf(c)

            If cVP._id = "" Then
                '' ur: war nur zu Testzwecken: 
                '' Call MsgBox("es ist noch kein VisboPortfolio angelegt")

                ' Portfolio-Name
                cVP.name = c.constellationName
                ' berechtiger User
                Dim user As New clsUser
                user.email = aktUser.email
                user.role = "Admin"
                cVP.users.Add(user)
                ' VisboCenter - Id
                cVP.vcid = aktVCid
                ' VisboProject-Type - Portfolio
                cVP.vpType = ptPRPFType.portfolio

                ' Erzeugen des VisboPortfolios in der Collection visboproject im akt. VisboCenter
                newVP = POSTOneVP(cVP)
                If newVP.Count > 0 Then
                    cVP._id = newVP.Item(0)._id
                Else
                    Throw New ArgumentException("FEHLER beim erstellen des VisboPortfolioProject")
                End If

            End If

            cVPf = clsConst2clsVPf(c)

            cVPf.vpid = cVP._id

            ' timestamp setzen

            cVPf.timestamp = DateTimeToISODate(Date.UtcNow)


            If cVP._id <> "" Then

                newVPf = POSTOneVPf(cVPf)

                If newVPf.Count > 0 Then
                    result = True
                End If

            End If
        Catch ex As Exception
            'Call MsgBox(ex.Message)
            Throw New ArgumentException(ex.Message)
        End Try

        storeConstellationToDB = result

    End Function

    ''' <summary>
    ''' Löschen des Portfolios  mit allen vorhandene Versionen aus der Datenbank
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Public Function removeConstellationFromDB(ByVal c As clsConstellation) As Boolean

        Dim result As Boolean = False

        Try
            Dim vpType As Integer = ptPRPFType.portfolio
            Dim cVPf As New clsVPf
            Dim cVP As New clsVP
            Dim newVP As New List(Of clsVP)
            Dim newVPf As New SortedList(Of Date, clsVPf)

            ' angepasst: 20180914: korrigieren, wenn ReST-Server geändert wurde
            'cVP = GETvpid(c.constellationName, vpType:=2)
            cVP = GETvpid(c.constellationName, ptPRPFType.portfolio)

            newVPf = GETallVPf(cVP._id, Date.Now)

            'aktuell müssen zum löschen eines Portfolios alle PortfolioVersionen gelöscht werden
            If newVPf.Count > 0 Then

                If newVPf.Count = 1 Then
                    result = DELETEOneVPf(cVP._id, newVPf.ElementAt(0).Value._id)
                Else
                    Dim lv As Integer = 0
                    Dim ok As Boolean = True
                    result = ok
                    While result And (lv < newVPf.Count)
                        lv = lv + 1
                        ok = DELETEOneVPf(cVP._id, newVPf.ElementAt(lv - 1).Value._id)
                        If lv = 1 Then
                            result = ok
                        Else
                            result = result And ok
                        End If
                    End While
                    'Call MsgBox("Es gab mehrer Portfolio-Versionen zu: " & c.constellationName)
                End If
            Else
                ' aktuell existiert keine PortfolioVersion zu vpid
                ' TODO: was ist, wenn nur der Token is dead war?!?!?
            End If

            If result = True Then
                result = DELETEOneVP(cVP._id)
            End If
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        removeConstellationFromDB = result

    End Function

    ''' <summary>
    ''' holt von allen Projekt-Varianten in AlleProjekte die Write-Protections
    ''' </summary>
    ''' <param name="AlleProjekte"></param>
    ''' <returns></returns>
    Public Function retrieveWriteProtectionsFromDB(ByVal AlleProjekte As clsProjekteAlle) As SortedList(Of String, clsWriteProtectionItem)

        Dim result As New SortedList(Of String, clsWriteProtectionItem)

        Try
            For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

                Dim wpItem As New clsWriteProtectionItem
                wpItem.pvName = kvp.Key
                Dim pname As String = Projekte.getPnameFromKey(kvp.Key)
                Dim vname As String = Projekte.getVariantnameFromKey(kvp.Key)
                wpItem = getWriteProtection(pname, vname, ptPRPFType.project)

                If Not result.ContainsKey(wpItem.pvName) Then
                    result.Add(wpItem.pvName, wpItem)
                End If

            Next
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try


        retrieveWriteProtectionsFromDB = result
    End Function


    ''' <summary>
    ''' löst von allen Projekt-Varianten des Users user die nonpermanent writeProtections
    ''' </summary>
    ''' <param name="user"></param>
    ''' <returns></returns>
    Public Function cancelWriteProtections(ByVal user As String) As Boolean

        Dim result As Boolean = False
        Dim vplist As New SortedList(Of String, clsVP)

        Try
            ' alle vp des aktuellen Users und aktuellen vc holen
            If VRScache.VPsN.Count > 0 Then
                vplist = VRScache.VPsN
            Else
                vplist = GETallVP(aktVCid, Nothing)
            End If

            For Each kvp As KeyValuePair(Of String, clsVP) In vplist

                If kvp.Value.lock.Count > 0 Then

                    ' holt zu der vpid die Varianten aus vpv Collection
                    Dim variantToProj As List(Of clsProjektWebShort) = GETallVPvShort(kvp.Value._id, , "", False,  , Date.Now)

                    ' Lock löschen für jede Variante des Projektes mit vpid
                    For Each vTp As clsProjektWebShort In variantToProj

                        result = result And DELETEVPLock(kvp.Value._id, vTp.variantName)

                    Next
                End If

            Next

        Catch ex As Exception
            Throw New ArgumentException(ex.Message & "Fehler in cancelWriteProtections:")
        End Try

        cancelWriteProtections = result
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
            Throw New ArgumentException(ex.Message)
        End Try
        retrieveRolesFromDB = result

    End Function

    ''' <summary>
    ''' speichert eine Rolle in der Datenbank; 
    ''' wenn insertNewDate = true: speichere eine neue Timestamp-Instanz 
    ''' andernfalls wird die Rolle Replaced 
    ''' </summary>
    ''' <param name="roleDef"></param>
    ''' <param name="insertNewDate"></param>
    ''' <param name="ts"></param>
    ''' <returns></returns>
    Public Function storeRoleDefinitionToDB(ByVal roleDef As clsRollenDefinition, ByVal insertNewDate As Boolean, ByVal ts As DateTime) As Boolean

        Dim result As Boolean = False

        Try
            Dim timestamp As String = DateTimeToISODate(ts.ToUniversalTime())

            Dim role As New clsVCrole
            role.copyFrom(roleDef)
            role.timestamp = timestamp

            If insertNewDate Then
                result = POSTOneVCrole(aktVCid, role)
            Else
                If VRScache.VCrole.ContainsKey(role.name) Then
                    role._id = VRScache.VCrole(role.name)._id
                    result = PUTOneVCrole(aktVCid, role)
                End If

                If result = False Then ' Rolle ist noch nicht vorhanden im VisboCenter, also neu erzeugen
                    result = POSTOneVCrole(aktVCid, role)
                End If
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        storeRoleDefinitionToDB = result
    End Function



    ''' <summary>
    '''  speichert eine Kostenart In der Datenbank; 
    '''  wenn insertNewDate = True: speichere eine neue Timestamp-Instanz 
    '''  andernfalls wird die Kostenart Replaced, sofern sie sich geändert hat  
    ''' </summary>
    ''' <param name="costDef"></param>
    ''' <param name="insertNewDate"></param>
    ''' <param name="ts"></param>
    ''' <returns></returns>
    Public Function storeCostDefinitionToDB(ByVal costDef As clsKostenartDefinition, ByVal insertNewDate As Boolean, ByVal ts As DateTime) As Boolean

        Dim result As Boolean = False

        Try
            Dim timestamp As String = DateTimeToISODate(ts.ToUniversalTime())

            Dim cost As New clsVCcost
            cost.copyFrom(costDef)
            cost.timestamp = timestamp

            If insertNewDate Then
                result = POSTOneVCcost(aktVCid, cost)
            Else

                If VRScache.VCcost.ContainsKey(cost.name) Then
                    cost._id = VRScache.VCcost(cost.name)._id
                    result = PUTOneVCcost(aktVCid, cost)
                End If

                If result = False Then  ' Kostenart ist noch nicht vorhanden im VisboCenter, also neu erzeugen
                    result = POSTOneVCcost(aktVCid, cost)
                End If
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try


        storeCostDefinitionToDB = result

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
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveCostsFromDB = result

    End Function

    ' ------------------------------------------------------------------------------------------
    '  Interne Funktionen für VisboRestServer - zugriff
    ' --------------------------------------------------------------------------------------------


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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="httpresp"></param>
    ''' <returns></returns>
    Private Function ReadResponseContent(ByRef httpresp As HttpWebResponse) As String
        'Private Function ReadResponseContent(ByRef resp As HttpWebResponse) As String
        Dim result As String = ""
        Try

            If IsNothing(httpresp) Then
                Throw New ArgumentNullException("HttpWebResponse ist Nothing")
            Else
                Dim statcode As HttpStatusCode = httpresp.StatusCode

                Try
                    Using sr As New StreamReader(httpresp.GetResponseStream)

                        result = sr.ReadToEnd()

                    End Using

                Catch ex As Exception

                End Try

            End If

        Catch ex As Exception
            Throw New ArgumentException("ReadResponseContent:" & ex.Message)
        End Try

        Return result

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
        ''ReDim bytes(1028)
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
    ''' holt zu dem Projekt projectName die zugehörige vpid vom Server
    ''' vpType = 1 ist Projekt; vpType = 0 ist Template (noch nicht fertig programmiert- ur:2018.07.24)
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <returns></returns>
    Private Function GETvpid(ByVal projectName As String, Optional ByVal vpType As Integer = ptPRPFType.project) As clsVP

        Dim vpid As String = ""
        Dim aktvp As New clsVP

        Try
            ' Alle VisboProjects über Server von WebServer/DB holen
            Dim anzLoop As Integer = 0
            'Dim allVP As New List(Of clsVP)
            While (vpid = "" And anzLoop <= 2)

                If VRScache.VPsN.Count > 0 Then
                    ' Id zu angegebenen Projekt herausfinden
                    If VRScache.VPsN.ContainsKey(projectName) Then
                        Dim vp As clsVP = VRScache.VPsN.Item(projectName)
                        vpid = vp._id
                        aktvp = vp
                    Else
                        VRScache.VPsN = GETallVP(aktVCid, vpType)
                    End If
                Else
                    VRScache.VPsN = GETallVP(aktVCid, ptPRPFType.all)
                End If

                anzLoop = anzLoop + 1
            End While

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        GETvpid = aktvp

    End Function

    ''' <summary>
    ''' holt zu dem Projekt/Portfolio mit der Id vpid den zugehörigen Projekt/Portfolio-Namen vom Server
    ''' </summary>
    ''' <param name="vpid"></param>
    ''' <returns></returns>
    Private Function GETpName(ByVal vpid As String) As String

        Dim pName As String = ""


        Try
            ' Alle VisboProjects über Server von WebServer/DB holen
            Dim anzLoop As Integer = 0

            If vpid <> "" Then

                While (pName = "" And anzLoop <= 2)

                    If VRScache.VPsId.ContainsKey(vpid) Then
                        ' pName zu angegebene vpid herausfinden
                        pName = VRScache.VPsId(vpid).name
                    Else
                        VRScache.VPsN = GETallVP(aktVCid, ptPRPFType.all)

                        Try
                            pName = VRScache.VPsId(vpid).name
                        Catch ex As Exception
                            pName = ""
                        End Try

                    End If

                    anzLoop = anzLoop + 1
                End While
            Else
                Throw New ArgumentException("Fehler in GETpName: keine vpid übergeben")
            End If
        Catch ex As Exception
            pName = ""
        End Try

        GETpName = pName

    End Function
    ''' <summary>
    ''' holt zu dem Projekt/Portfolio mit der Id vpid den zugehörigen Projekt/Portfolio-Namen vom Server
    ''' </summary>
    ''' <param name="vpid"></param>
    ''' <returns></returns>
    Private Function GETvpType(ByVal vpid As String) As Integer

        Dim vpType As Integer = ptPRPFType.all

        Try
            ' Alle VisboProjects über Server von WebServer/DB holen
            Dim anzLoop As Integer = 0

            If vpid <> "" Then

                While (vpType = ptPRPFType.all And anzLoop <= 2)

                    If VRScache.VPsId.ContainsKey(vpid) Then
                        ' pName zu angegebene vpid herausfinden
                        vpType = VRScache.VPsId(vpid).vpType
                    Else
                        VRScache.VPsN = GETallVP(aktVCid, ptPRPFType.all)

                        Try
                            vpType = VRScache.VPsId(vpid).vpType
                        Catch ex As Exception
                            vpType = ptPRPFType.all
                        End Try

                    End If

                    anzLoop = anzLoop + 1
                End While
            Else
                Throw New ArgumentException("Fehler in GETvpType: keine vpid übergeben")
            End If
        Catch ex As Exception
            vpType = ptPRPFType.all
        End Try

        GETvpType = vpType

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

                ' noch kein vcName gefunden
                If vcid = "" Then
                    VCs = GETallVC("")
                End If

                anzLoop = anzLoop + 1
            End While

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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
        Dim errmsg As String = ""
        Dim errcode As Integer
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
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCantwort = JsonConvert.DeserializeObject(Of clsWebVC)(Antwort)
            End Using

            If errcode = 200 Then           'success
                ' Call MsgBox(webVCantwort.message & vbCrLf & "es existieren " & webVCantwort.vc.Count & "VisboCenters")
                result = webVCantwort.vc
            Else

                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("GETallVC", errcode, errmsg & " : " & webVCantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        GETallVC = result

    End Function


    ''' <summary>
    ''' Holt alle VisboProject zu dem VisboCenter vcid 
    ''' und baut im Cache die Liste VPsId sortiert nach id und die VPsN sortiert nach Namen auf
    ''' </summary>
    ''' <param name="vcid">vcid = "": es werden alle VisboProjects dieses Users geholt
    '''                    sonst die visboProjects vom Visbocenter vcid</param>
    ''' <returns>nach Projektnamen sortierte Liste der VisboProjects</returns>
    Private Function GETallVP(ByVal vcid As String, Optional ByVal vptype As Integer = ptPRPFType.project) As SortedList(Of String, clsVP)

        Dim result As New SortedList(Of String, clsVP)          ' sortiert nach pname
        Dim secondResult As New SortedList(Of String, clsVP)    ' sortiert nach vpid
        Dim errmsg As String = ""
        Dim errcode As Integer
        Try
            Dim serverUriString As String
            Dim typeRequest As String = "/vp"

            ' URL zusammensetzen
            If vcid = "" Then
                serverUriString = serverUriName & typeRequest

                'If vptype <> ptPRPFType.portfolio Then
                If vptype <> ptPRPFType.project And vptype <> ptPRPFType.portfolio Then

                    '' kein bestimmter vp-Type gefragt
                Else
                    serverUriString = serverUriString & "?vpType=" & vptype.ToString
                End If

            Else
                serverUriString = serverUriName & typeRequest & "?vcid=" & vcid

                'If vptype <> ptPRPFType.portfolio Then
                If vptype <> ptPRPFType.project And vptype <> ptPRPFType.portfolio Then

                    '' kein bestimmter vp-Type gefragt
                Else
                    serverUriString = serverUriString & "&vpType=" & vptype.ToString
                End If

            End If

            Dim serverUri As New Uri(serverUriString)

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim Antwort As String
            Dim webVPantwort As clsWebVP = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVPantwort = JsonConvert.DeserializeObject(Of clsWebVP)(Antwort)
            End Using

            If errcode = 200 Then

                Select Case vptype

                    Case ptPRPFType.portfolio

                        For Each vp In webVPantwort.vp
                            result.Add(vp.name, vp)
                        Next


                    Case ptPRPFType.project

                        ' die erhaltenen Projekte werden in einer sortierten Liste gecacht
                        For Each vp In webVPantwort.vp

                            result.Add(vp.name, vp)

                            ' VPs nach Id sortiert gecacht
                            If Not VRScache.VPsId.ContainsKey(vp._id) Then
                                VRScache.VPsId.Add(vp._id, vp)
                            Else
                                VRScache.VPsId.Remove(vp._id)
                                VRScache.VPsId.Add(vp._id, vp)
                            End If


                            ' Cache-Struktur aufbauen für vpv, sortiert nach vpid
                            If Not VRScache.VPvs.ContainsKey(vp._id) Then
                                Dim leereListe As New SortedList(Of String, clsVarTs)
                                VRScache.VPvs.Add(vp._id, leereListe)

                            End If

                        Next

                    Case ptPRPFType.projectTemplate


                    Case Else

                        ' die erhaltenen Projekte/Portfolio-Projekte werden in einer sortierten Liste gecacht
                        For Each vp In webVPantwort.vp

                            result.Add(vp.name, vp)

                            ' VPs nach Id sortiert gecacht
                            If Not VRScache.VPsId.ContainsKey(vp._id) Then
                                VRScache.VPsId.Add(vp._id, vp)
                            Else
                                VRScache.VPsId.Remove(vp._id)
                                VRScache.VPsId.Add(vp._id, vp)
                            End If


                            ' Cache-Struktur aufbauen für vpv, sortiert nach vpid
                            If Not VRScache.VPvs.ContainsKey(vp._id) Then
                                Dim leereListe As New SortedList(Of String, clsVarTs)
                                VRScache.VPvs.Add(vp._id, leereListe)

                            End If

                        Next


                End Select

                GETallVP = result

            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("GETallVP", errcode, errmsg & " : " & webVPantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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
                                   Optional vpvid As String = "",
                                   Optional status As String = "",
                                   Optional refNext As Boolean = False,
                                   Optional ByVal variantName As String = Nothing,
                                   Optional ByVal storedAtorBefore As Date = Nothing) As List(Of clsProjektWebShort)

        Dim nothingToDo As Boolean = True
        Dim result As New List(Of clsProjektWebShort)
        Dim errmsg As String = ""
        Dim errcode As Integer = 200

        If Not (refNext Or status <> "") Then

            Try
                ' hier wird gecheckt, ob die Timestamps für vpid und variantName bereits im Cache sind
                nothingToDo = VRScache.existsInCache(vpid, variantName, vpvid, False, storedAtorBefore)
            Catch ex As Exception
                Call MsgBox("Fehler in existsInCache - Short")
            End Try

        Else
            nothingToDo = False
        End If

        If nothingToDo Then

            ' es existieren zu dieser vpid  und variantenName vpvs mit timestamps
            ' diese werden hier in die result-liste gebracht
            For Each kvp As KeyValuePair(Of String, SortedList(Of String, clsVarTs)) In VRScache.VPvs

                Dim clsVarTs_vpid As String = kvp.Key
                Dim clsVarTs_value As SortedList(Of String, clsVarTs) = kvp.Value

                For Each kvp1 As KeyValuePair(Of String, clsVarTs) In clsVarTs_value

                    Dim vname As String = kvp1.Key
                    Dim varts_liste As SortedList(Of Date, clsProjektWebShort) = kvp1.Value.tsShort

                    Dim found As Boolean = False
                    Dim i As Integer = varts_liste.Count - 1

                    While Not found And i >= 0
                        Dim ts As Date = varts_liste.ElementAt(i).Key
                        Dim shortproj As clsProjektWebShort = varts_liste.ElementAt(i).Value

                        If storedAtorBefore > Date.MinValue Then
                            ' größte, das kleiner als storeAtorBefore ist, als result zurückgeben
                            If ts <= storedAtorBefore Then

                                result.Add(shortproj)
                                found = True
                            Else
                                ' ProjShort in der Liste ist aktuell das am nächsten bei storedAtorBefore
                            End If
                        Else
                            result.Add(shortproj)
                        End If
                        i = i - 1
                    End While

                    ' wenn eine Variante angegeben ist, so nimm nur diese
                    If Not IsNothing(variantName) Then
                        If vname = variantName Then
                            Exit For
                        End If
                    End If
                Next
                If clsVarTs_vpid = vpid Then
                    Exit For
                End If
            Next
        Else

            Try

                Dim typeRequest As String = "/vpv"
                Dim serverUriString As String = serverUriName & typeRequest

                If vpvid <> "" Then
                    serverUriString = serverUriString & "/" & vpvid
                Else

                    serverUriString = serverUriString & "?"
                    serverUriString = serverUriString & "vcid=" & aktVCid

                    Dim refDate As String = DateTimeToISODate(storedAtorBefore)

                    If vpid <> "" Or storedAtorBefore > Date.MinValue Or variantName <> Nothing Then

                        If vpid <> "" Then
                            serverUriString = serverUriString & "&vpid=" & vpid

                            If storedAtorBefore > Date.MinValue Then
                                serverUriString = serverUriString & "&refDate=" & refDate
                                If refNext Then
                                    serverUriString = serverUriString & "&refNext=" & refNext.ToString
                                End If
                            Else
                                If refNext Then
                                    serverUriString = serverUriString & "&refDate=" & refDate
                                    serverUriString = serverUriString & "&refNext=" & refNext.ToString
                                End If
                            End If
                            If status <> "" Then
                                serverUriString = serverUriString & "&status=" & status
                            End If
                            If variantName <> Nothing Then
                                serverUriString = serverUriString & "&variantName=" & variantName
                            End If
                        Else
                            If storedAtorBefore > Date.MinValue Then
                                serverUriString = serverUriString & "&refDate=" & refDate
                                If refNext Then
                                    serverUriString = serverUriString & "&refNext=" & refNext.ToString
                                End If


                                If status <> "" Then
                                    serverUriString = serverUriString & "&status=" & status
                                End If
                                If variantName <> Nothing Then
                                    serverUriString = serverUriString & "&variantName=" & variantName
                                End If
                            Else
                                serverUriString = serverUriString & "&refDate=" & refDate
                                If refNext Then
                                    serverUriString = serverUriString & "&refDate=" & refDate
                                    serverUriString = serverUriString & "&refNext=" & refNext.ToString
                                End If
                                If status <> "" Then
                                    serverUriString = serverUriString & "&status=" & status
                                End If
                                If variantName <> Nothing Then
                                    serverUriString = serverUriString & "&variantName=" & variantName
                                End If

                            End If
                        End If

                        '    If vpid <> "" Or storedAtorBefore > Date.MinValue Or variantName <> Nothing Then

                        '    If vpid <> "" Then
                        '        serverUriString = serverUriString & "&vpid=" & vpid

                        '        If storedAtorBefore > Date.MinValue Then
                        '            serverUriString = serverUriString & "&refDate=" & refDate
                        '        End If

                        '        If variantName <> Nothing Then
                        '            serverUriString = serverUriString & "&variantName=" & variantName
                        '        End If
                        '    Else
                        '        If storedAtorBefore > Date.MinValue Then
                        '            serverUriString = serverUriString & "&refDate=" & refDate
                        '            If variantName <> Nothing Then
                        '                serverUriString = serverUriString & "&variantName=" & variantName
                        '            End If
                        '        Else
                        '            If variantName <> Nothing Then
                        '                serverUriString = serverUriString & "&variantName=" & variantName
                        '            End If
                        '        End If

                        '    End If

                    End If

                End If


                Dim serverUri As New Uri(serverUriString)

                Dim datastr As String = ""
                Dim encoding As New System.Text.UTF8Encoding()
                Dim data As Byte() = encoding.GetBytes(datastr)

                Dim Antwort As String
                Dim webVPvAntwort As clsWebVPv
                Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                    Antwort = ReadResponseContent(httpresp)
                    errcode = CType(httpresp.StatusCode, Integer)
                    errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                    webVPvAntwort = JsonConvert.DeserializeObject(Of clsWebVPv)(Antwort)
                End Using

                'If webVPvAntwort.state = "success" Then
                If errcode = 200 Then

                    ' Call MsgBox(webVPvAntwort.message & vbCrLf & "aktueller User hat " & webVPvAntwort.vpv.Count & " VisboProjectsVersions")
                    result = webVPvAntwort.vpv

                    If storedAtorBefore <= Date.MinValue And Not refNext And Not status <> "" Then
                        ' nur dann soll der Cache gefüllt werden, damit auch wirklich alle aktuellen Timestamps enthalten sind
                        VRScache.createVPvShort(result, Date.Now.ToUniversalTime)
                    End If


                Else

                    ' Fehlerbehandlung je nach errcode
                    Dim statError As Boolean = errorHandling_withBreak("GETallVPvShort", errcode, errmsg & " : " & webVPvAntwort.message)

                End If

            Catch ex As Exception
                Throw New ArgumentException(ex.Message)
            End Try

        End If

        GETallVPvShort = result

    End Function


    ''' <summary>
    ''' holt zu einer vpid alle VisboProjectsVersionen, wenn ein VarianteName angegeben ist, werden alle Versionen dieser Variante geholt
    ''' bei gegebenen storedAtorBefore nur die neueste Version zu diesem Datum
    ''' </summary>
    ''' <param name="vpid"></param>
    ''' <param name="vpvid"></param>
    ''' <param name="variantName"></param>
    ''' <param name="storedAtorBefore"></param>
    ''' <returns></returns>
    Private Function GETallVPvLong(ByVal vpid As String,
                                   Optional vpvid As String = "",
                                   Optional status As String = "",
                                   Optional refNext As Boolean = False,
                                   Optional ByVal variantName As String = Nothing,
                                   Optional ByVal storedAtorBefore As Date = Nothing) As List(Of clsProjektWebLong)

        Dim result As New List(Of clsProjektWebLong)
        Dim nothingToDo As Boolean = True
        Dim errmsg As String = ""
        Dim errcode As Integer

        If Not (refNext Or status <> "") Then

            Try
                ' hier wird gecheckt, ob die Timestamps für vpid und variantName bereits im Cache sind
                nothingToDo = VRScache.existsInCache(vpid, variantName, vpvid, True, storedAtorBefore)
            Catch ex As Exception
                Call MsgBox("Fehler in existsInCache - Long")
            End Try
        Else
            nothingToDo = False
        End If


        If nothingToDo Then

            ' es existieren zu dieser vpid  und variantenName vpvs mit timestamps
            ' diese werden hier in die result-liste gebracht
            For Each kvp As KeyValuePair(Of String, SortedList(Of String, clsVarTs)) In VRScache.VPvs

                Dim clsVarTs_vpid As String = kvp.Key
                Dim clsVarTs_value As SortedList(Of String, clsVarTs) = kvp.Value

                For Each kvp1 As KeyValuePair(Of String, clsVarTs) In clsVarTs_value

                    Dim vname As String = kvp1.Key
                    Dim varts_liste As SortedList(Of Date, clsProjektWebLong) = kvp1.Value.tsLong

                    Dim found As Boolean = False
                    Dim i As Integer = varts_liste.Count - 1

                    While Not found And i >= 0
                        Dim ts As Date = varts_liste.ElementAt(i).Key
                        Dim longproj As clsProjektWebLong = varts_liste.ElementAt(i).Value

                        If storedAtorBefore > Date.MinValue Then
                            ' größte, das kleiner als storeAtorBefore ist, als result zurückgeben
                            If ts <= storedAtorBefore Then

                                result.Add(longproj)
                                found = True
                            Else
                                ' ProjShort in der Liste ist aktuell das am nächsten bei storedAtorBefore
                            End If
                        Else
                            result.Add(longproj)
                        End If
                        i = i - 1
                    End While

                    ' wenn eine Variante angegeben ist, so nimm nur diese
                    If Not IsNothing(variantName) Then
                        If vname = variantName Then
                            Exit For
                        End If
                    End If
                Next
                If clsVarTs_vpid = vpid Then
                    Exit For
                End If
            Next

            '' es existieren zu dieser vpid  und variantenName vpvs mit timestamps
            '' diese werden hier in die result-liste gebracht
            'For Each kvp As KeyValuePair(Of Date, clsProjektWebLong) In VRScache.VPvs(vpid)(variantName).tsLong
            '    If storedAtorBefore > Date.MinValue Then

            '        If kvp.Key <= storedAtorBefore Then
            '            result.Add(kvp.Value)
            '        End If
            '    Else
            '        result.Add(kvp.Value)
            '    End If

            'Next
        Else

            Try

                Dim typeRequest As String = "/vpv"
                Dim serverUriString As String = serverUriName & typeRequest

                If vpvid <> "" Then
                    serverUriString = serverUriString & "/" & vpvid
                Else

                    serverUriString = serverUriString & "?"
                    serverUriString = serverUriString & "vcid=" & aktVCid

                    'Dim refDate As String = DateTimeToISODate(storedAtorBefore.AddMinutes(1.0))
                    Dim refDate As String = DateTimeToISODate(storedAtorBefore)

                    If vpid <> "" Or storedAtorBefore > Date.MinValue Or variantName <> Nothing Then

                        If vpid <> "" Then
                            serverUriString = serverUriString & "&vpid=" & vpid

                            If storedAtorBefore > Date.MinValue Then
                                serverUriString = serverUriString & "&refDate=" & refDate
                                If refNext Then
                                    serverUriString = serverUriString & "&refNext=" & refNext.ToString
                                End If
                            Else
                                If refNext Then
                                    serverUriString = serverUriString & "&refDate=" & refDate
                                    serverUriString = serverUriString & "&refNext=" & refNext.ToString
                                End If
                            End If
                            If status <> "" Then
                                serverUriString = serverUriString & "&status=" & status
                            End If
                            If variantName <> Nothing Then
                                serverUriString = serverUriString & "&variantName=" & variantName
                            End If
                        Else
                            If storedAtorBefore > Date.MinValue Then
                                serverUriString = serverUriString & "&refDate=" & refDate
                                If refNext Then
                                    serverUriString = serverUriString & "&refNext=" & refNext.ToString
                                End If


                                If status <> "" Then
                                    serverUriString = serverUriString & "&status=" & status
                                End If
                                If variantName <> Nothing Then
                                    serverUriString = serverUriString & "&variantName=" & variantName
                                End If
                            Else

                                If refNext Then
                                    serverUriString = serverUriString & "&refDate=" & refDate
                                    serverUriString = serverUriString & "&refNext=" & refNext.ToString
                                End If
                                If status <> "" Then
                                    serverUriString = serverUriString & "&status=" & status
                                End If
                                If variantName <> Nothing Then
                                    serverUriString = serverUriString & "&variantName=" & variantName
                                End If

                            End If
                        End If

                        ' es wird die Long-Version einer VisboProjectVersion angefordert
                        serverUriString = serverUriString & "&longList"

                    Else


                        ' Long-Version  angefordert
                        serverUriString = serverUriString & "&longList"


                    End If
                End If

                Dim serverUri As New Uri(serverUriString)

                Dim datastr As String = ""
                Dim encoding As New System.Text.UTF8Encoding()
                Dim data As Byte() = encoding.GetBytes(datastr)

                Dim Antwort As String
                Dim webVPvAntwort As clsWebLongVPv
                Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                    Antwort = ReadResponseContent(httpresp)
                    ' speichern von Error-Code und -Message für error-Handling
                    errcode = CType(httpresp.StatusCode, Integer)
                    errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                    webVPvAntwort = JsonConvert.DeserializeObject(Of clsWebLongVPv)(Antwort)
                End Using

                If errcode = 200 Then

                    result = webVPvAntwort.vpv

                    ' cache soll nur befüllt werden, wenn nicht explizit mit VisboProjectVersion-Id aufgerufen
                    If (vpvid = "") Then
                        ' nur dann soll der Cache gefüllt werden, damit auch wirklich alle aktuellen Timestamps enthalten sind
                        VRScache.createVPvLong(result, Date.Now.ToUniversalTime)
                    End If

                Else

                    ' Fehlerbehandlung je nach errcode
                    Dim statError As Boolean = errorHandling_withBreak("GETallVPvLong", errcode, errmsg & " : " & webVPvAntwort.message)

                End If

            Catch ex As Exception
                Throw New ArgumentException(ex.Message)
            End Try

        End If

        GETallVPvLong = result

    End Function


    ''' <summary>
    ''' Holt alle VisboProject-PortfolioVersionen zu dem aktuellen VISboCenter  und VisboProject-Id vpid
    ''' und baut im Cache die Liste VPsId sortiert nach id und die VPsN sortiert nach Namen auf
    ''' </summary>
    ''' <param name="vpid">vpid = "": es werden alle VisboportfolioVersions  dieser vpid geholt
    '''                    die jünger sind als timestamp</param>
    ''' <returns>nach Projektnamen sortierte Liste der VisboProjects</returns>
    Private Function GETallVPf(ByVal vpid As String, ByVal timestamp As Date) As SortedList(Of Date, clsVPf)

        Dim result As New SortedList(Of Date, clsVPf)          ' sortiert nach datum
        Dim secondResult As New SortedList(Of String, clsVPf)    ' sortiert nach vpid
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            Dim serverUriString As String
            Dim typeRequest As String = "/vp"

            ' URL zusammensetzen
            serverUriString = serverUriName & typeRequest & "/" & vpid & "/portfolio"
            Dim serverUri As New Uri(serverUriString)

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim Antwort As String
            Dim webVPfantwort As clsWebVPf = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVPfantwort = JsonConvert.DeserializeObject(Of clsWebVPf)(Antwort)
            End Using

            If errcode = 200 Then

                'die PortfolioVersionen werden nach Timestamp sortiert
                For Each vpf In webVPfantwort.vpf

                    Dim x As Date = CDate(vpf.timestamp)
                    Dim constellationName As String = GETpName(vpid)
                    If vpf.name = constellationName Then
                        result.Add(vpf.timestamp, vpf)
                    End If

                Next

                GETallVPf = result

            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("GETallVPf", errcode, errmsg & " : " & webVPfantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        GETallVPf = result

    End Function

    ''' <summary>
    ''' löscht eine VisboProjectVersion
    ''' </summary>
    ''' <param name="vpvid"></param>
    ''' <returns>true:  löschen erfolgreich
    '''          false: löschen hat nicht funktioniert</returns>
    Private Function DELETEOneVPv(ByVal vpvid As String) As Boolean

        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            ' URL zusammensetzen
            Dim typeRequest As String = "/vpv"
            Dim serverUriString As String = serverUriName & typeRequest

            If vpvid <> "" Then
                serverUriString = serverUriString & "/" & vpvid
            End If

            Dim serverUri As New Uri(serverUriString)

            ' DATA - Block zusammensetzen

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            ' Request absetzen
            Dim Antwort As String
            Dim webantwort As clsWebOutput = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "DELETE")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webantwort = JsonConvert.DeserializeObject(Of clsWebOutput)(Antwort)
            End Using

            If errcode = 200 Then
                result = True
            Else
                Dim statError As Boolean = errorHandling_withBreak("DELETEOneVPv", errcode, errmsg & " : " & webantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        DELETEOneVPv = result

    End Function

    ''' <summary>
    ''' löscht eine VisboPortfolioVersion
    ''' </summary>
    ''' <param name="vpid"></param>
    ''' <param name="vpfid"></param>
    ''' <returns>true:  löschen erfolgreich
    '''          false: löschen hat nicht funktioniert</returns>
    Private Function DELETEOneVPf(ByVal vpid As String, ByVal vpfid As String) As Boolean

        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            ' URL zusammensetzen
            Dim typeRequest As String = "/vp"
            Dim serverUriString As String = serverUriName & typeRequest

            If vpid <> "" And vpfid <> "" Then
                serverUriString = serverUriString & "/" & vpid & "/portfolio/" & vpfid
            End If

            Dim serverUri As New Uri(serverUriString)

            ' DATA - Block zusammensetzen

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            ' Request absetzen
            Dim Antwort As String
            Dim webantwort As clsWebVP = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "DELETE")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webantwort = JsonConvert.DeserializeObject(Of clsWebVP)(Antwort)
            End Using

            If errcode = 200 Then
                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("DELETEOneVPf", errcode, errmsg & " : " & webantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        DELETEOneVPf = result

    End Function

    ''' <summary>
    ''' Holt alle Rollen (vcrole) zu dem VisboCenter vcid
    ''' </summary>
    ''' <param name="vcid">vcid = "": es werden alle Rollen vom Visbocenter vcid  geholt</param>
    '''                    
    ''' <returns>Liste der Rollen</returns>
    Private Function GETallVCrole(ByVal vcid As String) As List(Of clsVCrole)

        Dim result As New List(Of clsVCrole)
        Dim errmsg As String = ""
        Dim errcode As Integer

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
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCroleantwort = JsonConvert.DeserializeObject(Of clsWebVCrole)(Antwort)
            End Using

            If errcode = 200 Then

                result = webVCroleantwort.vcrole

                ' hier werden die Rollen im Cache angelegt.
                For Each vcrole As clsVCrole In result

                    If Not VRScache.VCrole.ContainsKey(vcrole.name) Then
                        VRScache.VCrole.Add(vcrole.name, vcrole)
                    End If

                Next

            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("GETallVCrole", errcode, errmsg & " : " & webVCroleantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        GETallVCrole = result

    End Function




    ''' <summary>
    ''' erzeugt die Rolle role im VisboCenter vcid
    ''' </summary>
    ''' <param name="vcid"></param>
    ''' <param name="role"></param>
    ''' <returns></returns>
    Private Function POSTOneVCrole(ByVal vcid As String, ByVal role As clsVCrole) As Boolean

        Dim result As Boolean
        Dim errmsg As String = ""
        Dim errcode As Integer

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
            Dim data As Byte() = serverInputDataJson(role, "")


            Dim Antwort As String
            Dim webVCroleantwort As clsWebVCrole = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCroleantwort = JsonConvert.DeserializeObject(Of clsWebVCrole)(Antwort)
            End Using

            If errcode = 200 Then
                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTOneVCrole", errcode, errmsg & " : " & webVCroleantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        POSTOneVCrole = result

    End Function



    ''' <summary>
    ''' ändert die Rolle role im VisboCenter vcid
    ''' </summary>
    ''' <param name="vcid"></param>
    ''' <param name="role"></param>
    ''' <returns></returns>
    Private Function PUTOneVCrole(ByVal vcid As String, ByVal role As clsVCrole) As Boolean

        Dim result As Boolean
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            Dim serverUriString As String
            Dim typeRequest As String = "/vc"

            ' URL zusammensetzen
            If vcid = "" Then
                serverUriString = serverUriName & typeRequest
            Else
                serverUriString = serverUriName & typeRequest & "/" & vcid
            End If
            serverUriString = serverUriString & "/role/" & role._id

            Dim serverUri As New Uri(serverUriString)
            Dim data As Byte() = serverInputDataJson(role, "")


            Dim Antwort As String
            Dim webVCroleantwort As clsWebVCrole = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "PUT")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCroleantwort = JsonConvert.DeserializeObject(Of clsWebVCrole)(Antwort)
            End Using

            If errcode = 200 Then

                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("PUTOneVCrole", errcode, errmsg & " : " & webVCroleantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        PUTOneVCrole = result

    End Function

    ''' <summary>
    ''' Holt alle Kostenarten (vccost) zu dem VisboCenter vcid
    ''' </summary>
    ''' <param name="vcid">vcid = "": es werden alle Kostenarten vom Visbocenter vcid geholt</param>
    ''' <returns>Liste der Kostenarten</returns>
    Private Function GETallVCcost(ByVal vcid As String) As List(Of clsVCcost)

        Dim result As New List(Of clsVCcost)
        Dim errmsg As String = ""
        Dim errcode As Integer

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
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCcostantwort = JsonConvert.DeserializeObject(Of clsWebVCcost)(Antwort)
            End Using

            If errcode = 200 Then

                result = webVCcostantwort.vccost
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("GETalVCcost", errcode, errmsg & " : " & webVCcostantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        GETallVCcost = result

    End Function


    ''' <summary>
    ''' erzeugt die Kostenart cost im VisboCenter vcid
    ''' </summary>
    ''' <param name="vcid"></param>
    ''' <param name="cost"></param>
    ''' <returns></returns>
    Private Function POSTOneVCcost(ByVal vcid As String, ByVal cost As clsVCcost) As Boolean

        Dim result As Boolean
        Dim errmsg As String = ""
        Dim errcode As Integer

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
            Dim data As Byte() = serverInputDataJson(cost, "")


            Dim Antwort As String
            Dim webVCcostantwort As clsWebVCcost = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCcostantwort = JsonConvert.DeserializeObject(Of clsWebVCcost)(Antwort)
            End Using

            If errcode = 200 Then
                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTOneVCcost", errcode, errmsg & " : " & webVCcostantwort.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        POSTOneVCcost = result

    End Function



    ''' <summary>
    ''' ändert die Kostenart cost im VisboCenter vcid
    ''' </summary>
    ''' <param name="vcid"></param>
    ''' <param name="cost"></param>
    ''' <returns></returns>
    Private Function PUTOneVCcost(ByVal vcid As String, ByVal cost As clsVCcost) As Boolean

        Dim result As Boolean
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            Dim serverUriString As String
            Dim typeRequest As String = "/vc"

            ' URL zusammensetzen
            If vcid = "" Then
                serverUriString = serverUriName & typeRequest
            Else
                serverUriString = serverUriName & typeRequest & "/" & vcid
            End If
            serverUriString = serverUriString & "/cost/" & cost._id

            Dim serverUri As New Uri(serverUriString)
            Dim data As Byte() = serverInputDataJson(cost, "")


            Dim Antwort As String
            Dim webVCcostantwort As clsWebVCcost = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "PUT")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCcostantwort = JsonConvert.DeserializeObject(Of clsWebVCcost)(Antwort)
            End Using

            If errcode = 200 Then

                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("PUTOneVCcost", errcode, errmsg & " : " & webVCcostantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        PUTOneVCcost = result

    End Function



    ''' <summary>
    ''' ändert ein VisboProject
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird ein VisboProject geändert. user muss die Rechte haben, das checkt der Server</param>
    ''' <returns>Liste der VisboProjects</returns>
    Private Function PUTOneVP(ByVal vpid As String, ByVal vp As clsVP) As List(Of clsVP)

        Dim result As New List(Of clsVP)
        Dim errmsg As String = ""
        Dim errcode As Integer

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

            Dim data As Byte() = serverInputDataJson(vp, "")

            Dim Antwort As String
            Dim webVPantwort As clsWebVP = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "PUT")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVPantwort = JsonConvert.DeserializeObject(Of clsWebVP)(Antwort)
            End Using

            If errcode = 200 Then

                result = webVPantwort.vp

            Else

                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("PUTOneVP", errcode, errmsg & " : " & webVPantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        PUTOneVP = result

    End Function



    ''' <summary>
    ''' löscht den Lock eines Projektes/variante
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird dass VisboProject vpid gelöscht. user muss die Rechte haben, das checkt der Server</param>
    ''' <returns>true: gelöscht
    '''          false: konnte nicht gelöscht werden</returns>
    Private Function DELETEOneVP(ByVal vpid As String) As Boolean

        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            ' URL zusammensetzen
            Dim typeRequest As String = "/vp"
            Dim serverUriString As String = serverUriName & typeRequest

            If vpid <> "" Then
                serverUriString = serverUriString & "/" & vpid
            End If

            Dim serverUri As New Uri(serverUriString)

            ' DATA - Block zusammensetzen

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)


            ' Request absetzen
            Dim Antwort As String
            Dim webVP As clsWebVP = Nothing

            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "DELETE")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVP = JsonConvert.DeserializeObject(Of clsWebVP)(Antwort)
            End Using

            If errcode = 200 Then

                Dim pname As String = GETpName(vpid)

                If VRScache.VPsId.ContainsKey(vpid) Then
                    VRScache.VPsId.Remove(vpid)
                End If

                If VRScache.VPsN.ContainsKey(pname) Then
                    VRScache.VPsN.Remove(pname)
                End If
                result = True
            Else

                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("DELETEOneVP", errcode, errmsg & " : " & webVP.message)

            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        DELETEOneVP = result

    End Function


    ''' <summary>
    ''' Lockt ein Projekt/variante
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird ein VisboProject geändert. user muss die Rechte haben, das checkt der Server</param>
    ''' <returns>Liste der VisboProjects</returns>
    Private Function POSTVPLock(ByVal vpid As String, ByVal variantName As String) As Boolean


        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            ' URL zusammensetzen
            Dim serverUriString As String = ""
            Dim typeRequest As String = "/vp"

            If vpid = "" Then
                Call MsgBox("Fehler beim POSTVPLock")
            Else
                serverUriString = serverUriName & typeRequest & "/" & vpid & "/lock"
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
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVPLockantwort = JsonConvert.DeserializeObject(Of clsWebVPlock)(Antwort)
            End Using

            If errcode = 200 Then

                Dim pname As String = GETpName(vpid)

                Dim newLock As clsVPLock = webVPLockantwort.lock.ElementAt(0)
                If VRScache.VPsId(vpid).lock.Count = 0 Then
                    VRScache.VPsId(vpid).lock.Add(newLock)
                    VRScache.VPsN(pname).lock.Add(newLock)
                Else
                    Dim variantNotFound As Boolean = True
                    ' suchen, ob bereits ein Lock für diese Variante besteht, der dann erneuert wird.
                    For Each lastlock As clsVPLock In VRScache.VPsId(vpid).lock
                        If lastlock.variantName = newLock.variantName Then
                            variantNotFound = False
                            If VRScache.VPsId(vpid).lock.Contains(lastlock) Then
                                VRScache.VPsId(vpid).lock.Remove(lastlock)
                                VRScache.VPsId(vpid).lock.Add(newLock)
                            End If
                            If VRScache.VPsN(pname).lock.Contains(lastlock) Then
                                VRScache.VPsN(pname).lock.Remove(lastlock)
                                VRScache.VPsN(pname).lock.Add(newLock)
                            End If
                            Exit For
                        End If
                    Next
                    If variantNotFound Then
                        VRScache.VPsId(vpid).lock.Add(newLock)
                        VRScache.VPsN(pname).lock.Add(newLock)
                    End If

                End If


                ' Lock wurde richtig durchgeführt, wenn auch die Anzahl Lock im Cache-Speicher übereinstimmt
                result = VRScache.VPsId(vpid).lock.Count = VRScache.VPsN(pname).lock.Count

            Else

                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTVPLock", errcode, errmsg & " : " & webVPLockantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            ' URL zusammensetzen
            Dim typeRequest As String = "/vp"
            Dim serverUriString As String = serverUriName & typeRequest

            If vpid = "" Then
                serverUriString = serverUriString & "/lock"
            Else
                serverUriString = serverUriString & "/" & vpid & "/lock"
            End If
            'If variantName <> "" Then
            serverUriString = serverUriString & "?variantName=" & variantName
            'End If



            Dim serverUri As New Uri(serverUriString)


            ' DATA - Block zusammensetzen

            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)


            ' Request absetzen
            Dim Antwort As String
            Dim webVPLockantwort As clsWebVPlock = Nothing

            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "DELETE")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVPLockantwort = JsonConvert.DeserializeObject(Of clsWebVPlock)(Antwort)
            End Using

            If errcode = 200 Then

                Dim pname As String = GETpName(vpid)

                Dim anzLock As Integer = webVPLockantwort.lock.Count
                If anzLock = 0 Then
                    VRScache.VPsId(vpid).lock.Clear()
                Else
                    VRScache.VPsId(vpid).lock = webVPLockantwort.lock
                    VRScache.VPsN(pname).lock = webVPLockantwort.lock
                End If
                ''For Each lastlock As clsVPLock In VRScache.VPsId(vpid).lock
                ''    If lastlock.variantName = variantName Then
                ''        If VRScache.VPsId(vpid).lock.Contains(lastlock) Then
                ''            VRScache.VPsId(vpid).lock.Remove(lastlock)
                ''        End If

                ''        Exit For
                ''    End If
                ''Next

                ' Lock wurde richtig durchgeführt, wenn auch die Anzahl Lock im Cache-Speicher übereinstimmt
                result = VRScache.VPsId(vpid).lock.Count = VRScache.VPsN(pname).lock.Count

            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("DELETEVPLock", errcode, errmsg & " : " & webVPLockantwort.message)
            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        DELETEVPLock = result

    End Function

    ''' <summary>
    ''' Erzeugt die Variante  variantName zu dem VisboProject vpid
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird ein VisboProject geändert. user muss die Rechte haben, das checkt der Server</param>
    ''' <param name="variantName"></param>
    ''' <returns></returns>
    Private Function POSTVPVariant(ByVal vpid As String, ByVal variantName As String) As Boolean


        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer

        Dim webVPVar As clsWebVPVariant
        Dim Data() As Byte

        Try

            Dim typeRequest As String = "/vp"
            Dim serverUriString As String = serverUriName & typeRequest & "/" & vpid & "/variant"
            Dim serverUri As New Uri(serverUriString)

            Dim var As New clsVPvariant
            var.variantName = variantName
            var.email = aktUser.email

            Data = serverInputDataJson(var, typeRequest)

            Dim Antwort As String
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, Data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVPVar = JsonConvert.DeserializeObject(Of clsWebVPVariant)(Antwort)
            End Using

            If errcode = 200 Then
                Try
                    ' Variante variantName in Cache mitaufnehmen
                    var = webVPVar.Variant.ElementAt(0)
                    If Not VRScache.VPsId(vpid).Variant.Contains(var) Then
                        VRScache.VPsId(vpid).Variant.Add(var)
                    End If

                Catch ex As Exception

                End Try
                result = True

            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTVPVariant", errcode, errmsg & " : " & webVPVar.message)
            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        POSTVPVariant = result

    End Function


    ''' <summary>
    ''' löscht die Variante variantName eines Projektes
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird die Variante des VisboProject vpid gelöscht. user muss die Rechte haben, das checkt der Server</param>
    ''' <returns>true: gelöscht
    '''          false: konnte nicht gelöscht werden</returns>
    Private Function DELETEVPVariant(ByVal vpid As String, Optional ByVal varID As String = "") As Boolean

        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            ' URL zusammensetzen
            Dim typeRequest As String = "/vp"
            Dim serverUriString As String = serverUriName & typeRequest

            If vpid = "" Then
                Call MsgBox("Fehler in DELETEVPVariant: keine vpid angegeben")
            Else
                serverUriString = serverUriString & "/" & vpid & "/variant/" & varID

                Dim serverUri As New Uri(serverUriString)

                ' DATA - Block zusammensetzen

                Dim datastr As String = ""
                Dim encoding As New System.Text.UTF8Encoding()
                Dim data As Byte() = encoding.GetBytes(datastr)


                ' Request absetzen
                Dim Antwort As String
                Dim webVPVarAntwort As clsWebVPVariant = Nothing

                Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "DELETE")
                    Antwort = ReadResponseContent(httpresp)
                    errcode = CType(httpresp.StatusCode, Integer)
                    errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                    webVPVarAntwort = JsonConvert.DeserializeObject(Of clsWebVPVariant)(Antwort)
                End Using

                If errcode = 200 Then

                    Dim anzvar As Integer = webVPVarAntwort.Variant.Count
                    If anzvar = 0 Then
                        VRScache.VPsId(vpid).Variant.Clear()
                    Else
                        VRScache.VPsId(vpid).Variant = webVPVarAntwort.Variant
                    End If

                    Dim pname As String = GETpName(vpid)
                    ' Lock wurde richtig durchgeführt, wenn auch die Anzahl Lock im Cache-Speicher übereinstimmt
                    result = VRScache.VPsId(vpid).Variant.Count = VRScache.VPsN(pname).Variant.Count

                Else
                    ' Fehlerbehandlung je nach errcode
                    Dim statError As Boolean = errorHandling_withBreak("DELETEVPVariant", errcode, errmsg & " : " & webVPVarAntwort.message)
                End If

            End If    ' ende von if vpid <> ""

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        DELETEVPVariant = result

    End Function


    ''' <summary>
    ''' legt ein VisboProject/VisboPortfolio an
    ''' </summary>
    ''' <param name="vp">hier sind alle Daten des Projektes/Portfolios enthalten</param>
    ''' <returns>Liste mit dem angelegten VisboProject/VisboPortfolio inkl. kreierter _Id</returns>
    Private Function POSTOneVP(ByVal vp As clsVP) As List(Of clsVP)

        Dim result As New List(Of clsVP)
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            Dim serverUriString As String = ""
            Dim typeRequest As String = "/vp"

            ' URL zusammensetzen
            serverUriString = serverUriName & typeRequest
            Dim serverUri As New Uri(serverUriString)

            Dim data As Byte() = serverInputDataJson(vp, "")

            Dim Antwort As String
            Dim webVPantwort As clsWebVP = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVPantwort = JsonConvert.DeserializeObject(Of clsWebVP)(Antwort)
            End Using

            If errcode = 200 Then

                result = webVPantwort.vp
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTOneVP", errcode, errmsg & " : " & webVPantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        POSTOneVP = result

    End Function

    ''' <summary>
    ''' legt ein VisboPortfolio-Version an
    ''' </summary>
    ''' <param name="vpf">hier sind alle Daten des Projektes/Portfolios enthalten</param>
    ''' <returns>Liste mit dem angelegten VisboProject/VisboPortfolio inkl. kreierter _Id</returns>
    Private Function POSTOneVPf(ByVal vpf As clsVPf) As List(Of clsVPf)

        Dim result As New List(Of clsVPf)
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            Dim serverUriString As String = ""
            Dim typeRequest As String = "/vp"


            ' URL zusammensetzen
            If vpf.vpid <> "" Then
                serverUriString = serverUriName & typeRequest & "/" & vpf.vpid & "/portfolio"
            Else
                Throw New ArgumentException(" vpid wurde für das Portfolio nicht angegeben")
            End If
            Dim serverUri As New Uri(serverUriString)

            Dim data As Byte() = serverInputDataJson(vpf, "")

            Dim Antwort As String
            Dim webVPfantwort As clsWebVPf = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVPfantwort = JsonConvert.DeserializeObject(Of clsWebVPf)(Antwort)
            End Using

            If errcode = 200 Then

                result = webVPfantwort.vpf
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTOneVPf", errcode, errmsg & " : " & webVPfantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        POSTOneVPf = result

    End Function

    Private Function POSTpwforgotten(ByVal ServerURL As String, ByVal databaseName As String, ByVal username As String) As Boolean

        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            Dim serverUriString As String = ""
            Dim typeRequest As String = "/token/user/pwforgotten"


            ' URL zusammensetzen
            serverUriName = ServerURL
            serverUriString = serverUriName & typeRequest
            Dim serverUri As New Uri(serverUriString)

            ' user-email in Struktur zum übergeben
            Dim user As New clsUserLoginSignup
            user.email = username

            Dim data As Byte() = serverInputDataJson(user, "")

            Dim Antwort As String
            Dim webantwort As clsWebOutput = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                'webantwort = JsonConvert.DeserializeObject(Of clsWeboutput)(Antwort)
            End Using

            If errcode = 200 Then

                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTpwforgotten", errcode, errmsg & " : " & webantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        POSTpwforgotten = result

    End Function


    ''' <summary>
    ''' Umwandlung einen Datum des Typs Date in einen ISO-Datums-String
    ''' </summary>
    ''' <param name="datumUhrzeit"></param>
    ''' <returns></returns>
    Private Function DateTimeToISODate(ByVal datumUhrzeit As Date) As String

        Dim ISODateandTime As String = Nothing
        Dim ISODate As String = ""
        Dim ISOTime As String = ""

        If datumUhrzeit >= Date.MinValue And datumUhrzeit <= Date.MaxValue Then
            ' DatumUhrzeit wird um 1 Sekunde erhöht, dass die 1000-stel keine Rolle spielen
            datumUhrzeit = datumUhrzeit.AddSeconds(1.0)
            ISODateandTime = datumUhrzeit.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        End If

        DateTimeToISODate = ISODateandTime

    End Function

    ''' <summary>
    ''' Kopieren des ReST-Server Portfolios vpf in das der DB-Version clsConstellation
    ''' </summary>
    ''' <param name="vpf"></param>
    ''' <returns></returns>
    Private Function clsVPf2clsConstellation(ByVal vpf As clsVPf) As clsConstellation
        Dim result As New clsConstellation
        Dim hConstItem As clsConstellationItem

        Try

            With result
                .constellationName = vpf.name
                ' Aufbau der Constellation.allitems
                For Each hvpfItem As clsVPfItem In vpf.allItems

                    hConstItem = New clsConstellationItem
                    hConstItem = clsVPfItem2clsConstItem(hvpfItem)

                    Dim pvname As String = calcProjektKey(hConstItem.projectName, hConstItem.variantName)
                    If Not .Liste.ContainsKey(pvname) Then
                        result.Liste.Add(pvname, hConstItem)
                    End If

                Next
                .sortCriteria = vpf.sortType
                Dim hsortliste As SortedList(Of String, String) = .sortListe(vpf.sortType)
            End With

        Catch ex As Exception
            result = Nothing
        End Try

        clsVPf2clsConstellation = result

    End Function

    ''' <summary>
    ''' Kopieren des Portfolio c in das Portfolio des ReST-Servers vom Tyü clsVPf
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Private Function clsConst2clsVPf(ByVal c As clsConstellation) As clsVPf

        Dim result As New clsVPf
        Try
            Dim hvpid As String = ""
            Dim vpfItem As New clsVPfItem

            With result
                .name = c.constellationName
                ._id = ""

                ' angepasst: 20180914: ReST-Server muss auf ptPRPFType-Enumeration angepasst werden
                '.vpid = GETvpid(c.constellationName, vpType:=2)._id

                .vpid = GETvpid(c.constellationName, vpType:=ptPRPFType.portfolio)._id

                .timestamp = DateTimeToISODate(Date.Now)

                .sortType = c.sortCriteria
                ' .sortlist aufbauen aus c.sortlist
                For Each kvp As KeyValuePair(Of String, String) In c.sortListe(result.sortType)
                    hvpid = GETvpid(kvp.Value)._id
                    If Not .sortList.Contains(hvpid) Then
                        .sortList.Add(hvpid)
                    End If
                Next
                ' .allitems liste aufbauen aus c.allitems
                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In c.Liste
                    vpfItem = clsConstItem2clsVPfItem(kvp.Value)
                    If Not result.allItems.Contains(vpfItem) Then
                        result.allItems.Add(vpfItem)
                    End If
                Next
            End With
        Catch ex As Exception

        End Try

        clsConst2clsVPf = result

    End Function

    ''' <summary>
    ''' Kopieren des vpfItem clsVPfItem in ein Element vom Typ clsConstellationItem 
    ''' wird vorallem bei den Portfolios (da anders als in ursprünglichen DB Version) benötigt
    ''' </summary>
    ''' <param name="vpfItem"></param>
    ''' <returns></returns>
    Private Function clsVPfItem2clsConstItem(ByVal vpfItem As clsVPfItem) As clsConstellationItem
        Dim result As New clsConstellationItem

        Try
            With result

                .projectName = GETpName(vpfItem.vpid)
                .variantName = vpfItem.variantName
                .start = vpfItem.start
                .show = vpfItem.show
                .zeile = vpfItem.zeile
                .reasonToExclude = vpfItem.reasonToExclude
                .reasonToInclude = vpfItem.reasonToInclude

            End With

        Catch ex As Exception
            result = Nothing
        End Try

        clsVPfItem2clsConstItem = result

    End Function

    ''' <summary>
    ''' Kopieren des clsConstellationItem cItem in ein Element vom Typ clsVPfItem
    ''' wird vorallem bei den Portfolios (da anders als in ursprünglichen DB Version) benötigt
    ''' </summary>
    ''' <param name="cItem"></param>
    ''' <returns></returns>
    Private Function clsConstItem2clsVPfItem(ByVal cItem As clsConstellationItem) As clsVPfItem
        Dim result As New clsVPfItem
        Try
            With result

                result.name = cItem.projectName
                result.vpid = GETvpid(cItem.projectName)._id
                result._id = ""
                result.projectName = cItem.projectName
                result.variantName = cItem.variantName
                result.start = cItem.start
                result.show = cItem.show
                result.zeile = cItem.zeile
                result.reasonToExclude = cItem.reasonToExclude
                result.reasonToInclude = cItem.reasonToInclude

            End With

        Catch ex As Exception
            result = Nothing
        End Try

        clsConstItem2clsVPfItem = result

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="RestCall">RestCall-Routine für besser Fehler-Lokalisation</param>
    ''' <param name="errcode">RestCall-Error 2xx, 3xx, 4xx, 5xx</param>
    ''' <param name="webAntwortMsg">Message</param>
    ''' <param name="withBreak"></param>
    ''' <returns></returns>
    Public Function errorHandling_withBreak(ByVal restCall As String, ByVal errcode As Integer,
                                            ByVal webAntwortMsg As String, Optional ByVal withBreak As Boolean = False) As Boolean

        Dim result As Boolean = False

        Try

            Select Case errcode

                Case 400        ' Bad Request

                    If awinSettings.visboDebug Then
                        Call MsgBox("Fehler in " & restCall & " : " & webAntwortMsg)
                    End If
                    If withBreak Then
                        Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)
                    End If


                Case 401        ' Unauthorized

                    token = ""
                    Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)

                Case 402        'Payment Required

                    If awinSettings.visboDebug Then
                        Call MsgBox("Fehler in " & restCall & " : " & webAntwortMsg)
                    End If
                    If withBreak Then
                        Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)
                    End If

                Case 403        ' Forbidden

                    'Call MsgBox("Fehler in GETallVPvShort: " & errmsg & " : " & webVPvAntwort.message)
                    Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)

                Case 404 To 408

                Case 409        ' Conflict

                    If awinSettings.visboDebug Then
                        Call MsgBox("Fehler in " & restCall & " : " & webAntwortMsg)
                    End If
                    If withBreak Then
                        Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)
                    End If

                Case 410 To 499

                    If awinSettings.visboDebug Then
                        Call MsgBox("Fehler in " & restCall & " : " & webAntwortMsg)
                    End If
                    If withBreak Then
                        Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)
                    End If

                Case 300 To 399

                    If awinSettings.visboDebug Then
                        Call MsgBox("Fehler in " & restCall & " : " & webAntwortMsg)
                    End If
                    'If withBreak Then
                    Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)
                    'End If

                Case 500 To 599     ' ServerIssue (internal Server Error)

                    Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)

                Case Else

            End Select

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        errorHandling_withBreak = result

    End Function
    '---------------------------------------------------------------------------------------------------------------
    '
    ' TODO: ur: Funktionen die für den Zugriff auf DB über ReST-Server noch fehlern
    '
    '---------------------------------------------------------------------------------------------------------------

    ''' <summary>
    '''  speichert einen Filter mit Namen 'name' in der Datenbank
    ''' </summary>
    ''' <param name="ptFilter"></param>
    ''' <param name="selfilter"></param>
    ''' <returns></returns>
    Public Function storeFilterToDB(ByVal ptFilter As clsFilter, ByRef selfilter As Boolean) As Boolean
        storeFilterToDB = True
    End Function



    ''' <summary>
    ''' Alle Abhängigkeiten aus der Datenbank lesen
    ''' und als Ergebnis ein Liste von Abhängigkeiten zurückgeben
    ''' </summary>
    ''' <returns></returns>
    Public Function retrieveDependenciesFromDB() As clsDependencies
        retrieveDependenciesFromDB = New clsDependencies
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
    ''' speichert Projekt-Dependencies in DB 
    ''' </summary>
    ''' <param name="d"></param>
    ''' <returns></returns>
    Public Function storeDependencyofPToDB(ByVal d As clsDependenciesOfP) As Boolean

        Dim result As Boolean = False
        storeDependencyofPToDB = result

    End Function
    ''    {

    ''        Try
    ''        {
    ''            var depDB = New clsDependenciesOfPDB();
    ''            depDB.copyFrom(d);
    ''            depDB.Id = depDB.projectName;

    ''            bool alreadyExisting = CollectionDependencies.AsQueryable < clsDependenciesOfPDB > ()
    ''.Any(p >= p.projectName == d.projectName);

    ''            If (alreadyExisting)
    ''            {
    ''                var filter = Builders < clsDependenciesOfPDB > .Filter.Eq("projectName", d.projectName);
    ''                var rResult = CollectionDependencies.ReplaceOne(filter, depDB);
    ''                If (rResult.ModifiedCount > 0)
    ''                {
    ''                    Return True;
    ''                }
    ''                Else
    ''                {
    ''                    Return False;
    ''                }
    ''            }
    ''            Else
    ''            {
    ''                CollectionDependencies.InsertOne(depDB);
    ''                Return True;
    ''            }

    ''        }
    ''        Catch (Exception)
    ''        {

    ''            Return False;
    ''        }


End Class

