
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
Imports System.Deployment.Application
Public Class Request

    'public serverUriName ="http://visbo.myhome-server.de:3484" 
    'Public serverUriName As String = "http://localhost:3484"

    Private serverUriName As String = ""

    Private version As System.Version
    Private visboContentType As String = "application/json"

    Private visboUserAgent As String = " (" & System.Environment.OSVersion.ToString & ";" & System.Environment.OSVersion.Platform.ToString & ")"

    Private aktVCid As String = ""

    Private token As String = ""
    Private VCs As New List(Of clsVC)

    Private VRScache As New clsCache
    ' hierin werden  alle Visbo-Projects und 
    ' die vom Server bereits angeforderten VisboProjectsVersionsgecacht
    '
    ' Private VPs As New SortedList(Of String, clsVP)
    '                                     vpid                  vname    timestamp-Liste, projectshort
    ' Private VPvCache As New SortedList(Of String, SortedList(Of String, clstest))
    ' Private VPvCache As New clsCache


    Private aktUser As clsUserReg = Nothing
    Private netcred As NetworkCredential

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
    Public Function login(ByVal ServerURL As String,
                          ByVal databaseName As String,
                          ByVal username As String,
                          ByVal dbPasswort As String,
                          ByRef err As clsErrorCodeMsg) As Boolean

        Dim typeRequest As String = "/token/user/login"
        'Dim typeRequest As String = "/token/user/signup"
        Dim serverUri As New Uri(ServerURL & typeRequest)
        Dim loginOK As Boolean = False
        Dim errcode As Integer = 0
        Dim errmsg As String = ""
        Dim httpresp_sav As HttpWebResponse

        Try
            If Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
                version =
                  Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion()
                visboUserAgent = visboClient & version.ToString & visboUserAgent
            Else
                ' Nicht via ClickOnce installiert
                visboUserAgent = visboClient & visboUserAgent
            End If

            Dim user As New clsUserLoginSignup
            user.email = LCase(username)
            user.password = dbPasswort
            'user.email = "markus.seyfried@visbo.de"
            'user.password = "visbo123"

            ' Konvertiere die erforderlichen Inputdaten des Requests vom Typ typeRequest (von der Struktur cls??) in ein Json-ByteArray
            Dim data() As Byte
            data = serverInputDataJson(user, typeRequest)


            Dim loginAntwort As New clsWebTokenUserLoginSignup
            Dim Antwort As String
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                httpresp_sav = httpresp     ' sichern der Server-Antwort
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription

            End Using

            If awinSettings.visboDebug Then
                Call MsgBox(loginAntwort.message)
            End If

            If errcode = 200 Then

                loginAntwort = JsonConvert.DeserializeObject(Of clsWebTokenUserLoginSignup)(Antwort)

                loginOK = True
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
                Else
                    ' alle vps dieses aktVCid lesen und im Cache speichern
                    Try
                        Dim err1 As New clsErrorCodeMsg
                        VRScache.VPsN = GETallVP(aktVCid, err1, ptPRPFType.all)
                    Catch ex As Exception

                    End Try

                End If

            Else
                token = ""
                serverUriName = ServerURL
                aktUser = Nothing
                If awinSettings.visboDebug Then
                    Call MsgBox("( " & CType(errcode, Integer).ToString & ") : " & errmsg & " : " & loginAntwort.message)
                End If

                err.errorCode = errcode
                err.errorMsg = "Login" & " : " & errmsg & " : " & loginAntwort.message

                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("Login", errcode, errmsg & " : " & loginAntwort.message)

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
    Public Function pwforgotten(ByVal ServerURL As String, ByVal databaseName As String, ByVal username As String, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Try
            result = POSTpwforgotten(ServerURL, databaseName, username, err)

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
    ''' <param name="err"></param>
    ''' <returns></returns>
    Public Function projectNameAlreadyExists(ByVal projectname As String, ByVal variantname As String,
                                             ByVal storedAtorBefore As DateTime,
                                             ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try
            If storedAtorBefore <= Date.MinValue Then
                storedAtorBefore = DateTime.Now.AddDays(1).ToUniversalTime()
            Else
                storedAtorBefore = storedAtorBefore.ToUniversalTime()
            End If

            Dim vpid As String = ""

            vpid = GETvpid(projectname, err)._id

            ' tk 28.12.18 hier eigentlich kritisch
            ' es kann nämlich den Fall geben, dass variantName <> "" existiert, aber variantNAme = "" existiert nicht 
            ' vorläufig dringelassen - es wird gecheckt, ob der Fall überhaupt auftreten kann bzw. ob das nicht grundsätzlich verhindert werden soll  
            'If vpid <> "" Then
            If vpid <> "" And variantname <> "" Then
                ' nachsehen, ob im VisboProject diese Variante zum Zeitpunkt storedAtorBefore bereits created war
                For Each vpVar As clsVPvariant In VRScache.VPsN(projectname).Variant

                    If vpVar.variantName = variantname Then
                        ' es muss mindestens eine VPV zu dieser Variante geben
                        If vpVar.vpvCount > 0 Then
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
    ''' <param name="err"></param>
    ''' </summary>
    ''' <returns>Collection, absteigend sortiert</returns>
    Public Function retrieveZeitstempelFromDB(ByRef err As clsErrorCodeMsg) As Collection

        Dim resultCollection As New Collection

        Try

            ' alle VisboProjectVersions vom Server anfordern
            ' ur:08.06.2018: wird in globale Variable gecacht: Dim allVPv As New List(Of clsProjektWebShort)

            Dim allVPv As New List(Of clsProjektWebShort)
            allVPv = GETallVPvShort("", err)

            ' alle vorhandenen Timestamps in der resultCollection sammeln
            Dim sl As New SortedList(Of Date, Date)
            For Each shortproj As clsProjektWebShort In allVPv
                If Not sl.ContainsKey(shortproj.timestamp) Then
                    sl.Add(shortproj.timestamp, shortproj.timestamp)
                End If
            Next

            For i As Integer = sl.Count - 1 To 0 Step -1
                Dim kvp As KeyValuePair(Of DateTime, DateTime) = sl.ElementAt(i)
                resultCollection.Add(kvp.Value.ToLocalTime())
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
    Public Function retrieveZeitstempelFromDB(ByVal pvName As String, ByRef err As clsErrorCodeMsg) As Collection

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
            vpid = GETvpid(projectName, err)._id

            If vpid <> "" Then
                ' gewünschte Variante vom Server anfordern
                Dim allVPv As New List(Of clsProjektWebShort)
                allVPv = GETallVPvShort(vpid:=vpid, err:=err,
                                        status:="", refNext:=False, variantName:=variantName,
                                        storedAtorBefore:=Nothing, fromReST:=False)

                ' alle vorhandenen Timestamps zu einem pvName in die ErgebnisCollection sammeln
                Dim sl As New SortedList(Of Date, Date)
                For Each shortproj As clsProjektWebShort In allVPv
                    If Not sl.ContainsKey(shortproj.timestamp) Then
                        sl.Add(shortproj.timestamp, shortproj.timestamp)
                    End If
                Next

                For i As Integer = sl.Count - 1 To 0 Step -1
                    Dim kvp As KeyValuePair(Of DateTime, DateTime) = sl.ElementAt(i)
                    ergebnisCollection.Add(kvp.Value.ToLocalTime)
                Next i

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveZeitstempelFromDB = ergebnisCollection

    End Function
    ''' <summary>
    ''' bringt für die angegebene Projekt-Variante den ersten und den letzten Zeitstempel  zurück 
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <returns>Collection, absteigend sortiert</returns>
    Public Function retrieveZeitstempelFirstLastFromDB(ByVal pvName As String, ByRef err As clsErrorCodeMsg) As Collection

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
            vpid = GETvpid(projectName, err)._id

            If vpid <> "" Then

                Dim hresultFirst As New List(Of clsProjektWebShort)

                hresultFirst = GETallVPvShort(vpid:=vpid, err:=err,
                                              vpvid:="",
                                              status:="", refNext:=True,
                                              variantName:=variantName,
                                              storedAtorBefore:=Nothing,
                                              fromReST:=False)


                Dim anzResult As Integer = hresultFirst.Count
                If anzResult >= 0 Then
                    ergebnisCollection.Add(hresultFirst.Item(anzResult - 1).timestamp.ToLocalTime)
                End If

                If err.errorCode = 200 Then
                    Dim hresultLast As New List(Of clsProjektWebShort)

                    hresultLast = GETallVPvShort(vpid:=vpid, err:=err,
                                                 status:="", refNext:=False,
                                                 variantName:=variantName,
                                                 storedAtorBefore:=Date.Now.ToUniversalTime,
                                                 fromReST:=False)

                    If hresultLast.Count >= 0 Then
                        ergebnisCollection.Add(hresultLast.Item(0).timestamp.ToLocalTime)
                    End If
                End If

            Else


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
                                               ByVal onlyLatest As Boolean, ByRef err As clsErrorCodeMsg) _
                                               As SortedList(Of String, clsProjekt)

        Dim result As New SortedList(Of String, clsProjekt)
        Dim diffRCBeginn As Date = Date.Now
        Dim diffRC As Long
        Dim diffCopy As Long

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
                Try

                    VRScache.VPsN = GETallVP(aktVCid, err, ptPRPFType.project)

                    If err.errorCode = 200 Then    ' Cache wurde erfolgreich gefüllt

                        Dim VisboPv_all As New List(Of clsProjektWebLong)
                        'VisboPv_all = GETallVPvLong("", err, , , , variantName, aktDate)
                        VisboPv_all = GETallVPvLong(vpid:="", err:=err, variantName:=variantName, storedAtorBefore:=aktDate)

                        'ur: nur zu testzwecken eingefügt
                        'diffRC = DateDiff(DateInterval.Second, diffRCBeginn, Date.Now)

                        'Dim copyBeginn As Date = Date.Now

                        For Each webProj As clsProjektWebLong In VisboPv_all

                            If (webProj.startDate <= zeitraumEnde And
                                    webProj.endDate >= zeitraumStart And
                                    webProj.timestamp <= storedLatest) Then

                                hproj = New clsProjekt
                                Dim vp As clsVP = Nothing

                                If VRScache.VPsN.ContainsKey(webProj.name) Then

                                    ' vp zum webProj aus dem Cache holen (keine Portfolios im Cache)
                                    vp = VRScache.VPsN(webProj.name)
                                    webProj.copyto(hproj, vp)

                                    Dim a As Integer = hproj.dauerInDays
                                    Dim key As String = Projekte.calcProjektKey(hproj)
                                    If Not result.ContainsKey(key) Then
                                        result.Add(key, hproj)
                                    End If
                                Else
                                    ' webProj war ein Portfolio-Projekt
                                    ' und wird übergangen
                                End If

                            End If

                        Next

                        'ur: nur zu testzwecken eingefügt
                        'diffCopy = DateDiff(DateInterval.Second, copyBeginn, Date.Now)

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
                        If awinSettings.englishLanguage Then
                            Call MsgBox("Error, while reading the Projects from DB: " & err.errorMsg)
                        Else
                            Call MsgBox("Fehler beim Lesen der Projekte aus dem DB: " & err.errorMsg)
                        End If

                    End If

                Catch ex As Exception

                End Try

            Else
                Try

                    '  Projekt angegeben: d.h. es werden alle Timestamps der übergebenen Projekt-Variante zurückgegeben
                    Dim vp As clsVP = GETvpid(projectname, err)
                    Dim vpid As String = vp._id
                    If vpid <> "" Then
                        ' gewünschten Varianten vom Server anfordern
                        Dim allVPv As New List(Of clsProjektWebLong)
                        allVPv = GETallVPvLong(vpid, err, , , , variantName, storedLatest)

                        For Each webProj As clsProjektWebLong In allVPv
                            If webProj.timestamp >= storedEarliest Then

                                hproj = New clsProjekt

                                webProj.copyto(hproj, vp)
                                Dim a As Integer = hproj.dauerInDays
                                Dim key As String = Projekte.calcProjektKey(hproj)
                                If Not result.ContainsKey(key) Then
                                    result.Add(key, hproj)
                                End If

                            End If

                        Next

                    End If

                Catch ex As Exception

                End Try

            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveProjectsFromDB = result

        If awinSettings.visboDebug Then
            Call MsgBox("RestTime: " & diffRC.ToString & vbLf & "CopyTime: " & diffCopy.ToString)
        End If


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
                                             ByVal storedAtOrBefore As DateTime,
                                             ByRef err As clsErrorCodeMsg) As clsProjekt
        Dim result As clsProjekt = Nothing

        storedAtOrBefore = storedAtOrBefore.ToUniversalTime

        Try
            Dim hproj As New clsProjekt
            Dim vpid As String = ""
            'ur:24.02.19: 
            'Dim vp As clsVP = GETvpid(projectname, err)

            ' projectname aus allen vps heraussuchen
            Dim vp As clsVP = GETvpid(projectname, err, ptPRPFType.all)

            vpid = vp._id

            If vpid <> "" Then
                ' gewünschte Variante vom Server anfordern
                Dim allVPv As New List(Of clsProjektWebLong)
                'allVPv = GETallVPvLong(vpid, err, , , , variantname, storedAtOrBefore)
                allVPv = GETallVPvLong(vpid:=vpid,
                                       err:=err,
                                       vpvid:="",
                                       status:="",
                                       refNext:=False,
                                       variantName:=variantname,
                                       storedAtorBefore:=storedAtOrBefore)

                If allVPv.Count > 0 Then
                    Dim webProj As clsProjektWebLong = allVPv.ElementAt(0)
                    webProj.copyto(hproj, vp)

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
    Public Function renameProjectsInDB(ByVal oldName As String,
                                       ByVal newName As String,
                                       ByVal userName As String,
                                       ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Try
            If projectNameAlreadyExists(newName, "", DateTime.Now, err) Then

                renameProjectsInDB = result

            Else

                Dim chkOk As Boolean = True

                ' hier wird überprüft, ob das Projekt selbst
                ' und auch keine der Varianten von einem anderen User schreibgeschützt ist

                chkOk = checkChgPermission(oldName, "", userName, err)

                If chkOk Then

                    Dim vp As New clsVP
                    Dim vpList As New List(Of clsVP)

                    Dim vpid As String = GETvpid(oldName, err)._id
                    If vpid <> "" Then

                        If VRScache.VPsN.ContainsKey(oldName) Then

                            vp = VRScache.VPsN(oldName)
                            vp.name = newName

                            vpList = PUTOneVP(vpid, vp, err)
                            ' rename war korrekt, wenn in vplist ein und zwar nur ein VisboProject zurückgegeben wurde.
                            If vpList.Count = 1 Then
                                If VRScache.VPsN.Remove(oldName) Then
                                    vp._id = vpid
                                    vp.name = newName
                                    VRScache.VPsN.Add(newName, vp)
                                End If

                            End If

                        End If

                    End If

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
    Public Function storeProjectToDB(ByVal projekt As clsProjekt,
                                     ByVal userName As String,
                                     ByRef mergedProj As clsProjekt,
                                     ByRef err As clsErrorCodeMsg,
                                     Optional ByVal attrToStore As Boolean = False) As Boolean

        Dim result As Boolean = False
        Dim errmsg As String = ""
        'Dim errcode As Integer

        Try

            Dim webVP As New clsWebVP
            Dim vpErg As New List(Of clsVP)
            'Dim data() As Byte

            Dim pname As String = projekt.name
            Dim vname As String = projekt.variantName
            Dim standardVariante As String = ""

            Dim aktvp As clsVP = GETvpid(pname, err, projekt.projectType)
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

                If Not IsNothing(projekt.kundenNummer) Then
                    VP.kundennummer = projekt.kundenNummer
                Else
                    VP.kundennummer = ""
                End If


                vpErg = POSTOneVP(VP, err)


                If vpErg.Count > 0 Then

                    ' vpErg.ElementAt(0) ist nun das aktuelle VP
                    vpid = vpErg.ElementAt(0)._id
                    aktvp = vpErg.ElementAt(0)
                    storedVP = (vpid <> "")

                    ' VP im Cache ergänzen
                    If VRScache.VPsN.ContainsKey(aktvp.name) Then
                        VRScache.VPsN.Remove(aktvp.name)
                        VRScache.VPsN.Add(aktvp.name, aktvp)
                    Else
                        VRScache.VPsN.Add(aktvp.name, aktvp)
                    End If
                    If VRScache.VPsId.ContainsKey(vpid) Then
                        VRScache.VPsId.Remove(vpid)
                        VRScache.VPsId.Add(vpid, aktvp)
                    Else
                        VRScache.VPsId.Add(vpid, aktvp)
                    End If

                Else
                    Throw New ArgumentException(err.errorCode & vbLf & "Das VisboProject existiert nicht und konnte auch nicht erzeugt werden!")
                End If


                ' hier wird der Fall behandelt : Anlegen einer Basis-Variante-Version, wenn der aktuelle varianteNAme <> "" ist

                '--------------------------------------------------------
                '     Basis-Variante erzeugen aus gegebener Variante
                '--------------------------------------------------------

                projekt.variantName = standardVariante ' STANDARD-Variante

                ' schreiben der Basis Variante 
                Dim erfolgreich As Boolean = POSTOneVPv(vpid, projekt, userName, err)

                If erfolgreich Then
                    projekt.variantName = vname
                Else

                End If


            Else
                Try
                    ' KundenNummer in vorhandenem VP ergänzen

                    If (aktvp.kundennummer = "") And (projekt.kundenNummer <> "") Then

                        aktvp.kundennummer = projekt.kundenNummer
                        Dim vpList As List(Of clsVP) = PUTOneVP(vpid, aktvp, err)

                    Else

                        If attrToStore Then

                            If String.Compare(aktvp.kundennummer, projekt.kundenNummer) = 0 Then
                                aktvp.kundennummer = projekt.kundenNummer
                                Dim vpList As List(Of clsVP) = PUTOneVP(vpid, aktvp, err)
                            End If

                        End If
                        ' nothing to do
                    End If

                Catch ex As Exception
                    Call MsgBox("Fehler beim Update von VP")
                End Try

            End If      ' Ende von "If Not storedVP Then"

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
                    storedVPVariant = POSTVPVariant(vpid, vname, err)
                Else
                    ' zu diesem Projekt gibt es nur die Standardvariante = > nichts tun
                    storedVPVariant = True
                End If
            End If

            ' Projekt ist bereits in VisboProjects Collection gespeichert, es existiert eine vpid
            If storedVP And storedVPVariant Then

                '--------------------------------------------------------
                '' ' jetzt muss noch VisboProjectVersion gespeichert werden
                '
                '     variantName-Variante erzeugen 
                '--------------------------------------------------------
                projekt.variantName = vname
                result = POSTOneVPv(vpid, projekt, userName, err)

                ' hier wird behandelt, wenn  von Seiten der RessourceManager konkurrierendes Schreiben vorkommt.
                If result = False Then
                    Select Case err.errorCode

                        Case 409

                            If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Then
                                Dim errNew As New clsErrorCodeMsg
                                Dim newResult As Boolean = result
                                Dim loopIndex As Integer = 1
                                While (newResult = False) And (loopIndex <= 10)

                                    Dim summaryRoleIDs As New Collection
                                    summaryRoleIDs.Add(myCustomUserRole.specifics)

                                    Dim newproj As clsProjekt = retrieveOneProjectfromDB(projekt.name, projekt.variantName, Date.Now, errNew)

                                    If Not IsNothing(newproj) Then
                                        If Not newproj.isIdenticalTo(projekt) Then
                                            ' Merge der geänderten Ressourcen => neues Projekt "mergeProj"
                                            mergedProj = newproj.deleteAndMerge(summaryRoleIDs, Nothing, projekt)
                                            newResult = POSTOneVPv(vpid, mergedProj, userName, err)
                                            If Not newResult Then
                                                mergedProj = Nothing
                                            End If
                                        End If
                                    Else
                                        err = errNew
                                    End If
                                    loopIndex = loopIndex + 1

                                End While

                                result = newResult

                            End If
                        Case Else
                            ' nothing to do
                    End Select

                End If

            End If

        Catch ex As Exception
            'Throw New ArgumentException(ex.Message & ": storeProjectToDB")
        End Try

        storeProjectToDB = result

    End Function



    ''' <summary>
    ''' liefert alle Varianten Namen eines bestimmten Projektes zurück 
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <returns></returns>
    Public Function retrieveVariantNamesFromDB(ByVal projectName As String, ByRef err As clsErrorCodeMsg) As Collection

        Dim ergebnisCollection As New Collection

        Try
            Dim vpid As String = ""

            ' nun ist sicher die VPs aufgebaut
            Dim vp As clsVP = GETvpid(projectName, err)

            If vp._id <> "" Then
                ' alle Variantenamen in der Collection sammeln
                For Each vpVar As clsVPvariant In vp.Variant
                    ergebnisCollection.Add(vpVar.variantName, vpVar.variantName)
                Next
            Else

            End If

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
                                                          ByVal storedAtOrBefore As DateTime,
                                                          ByRef err As clsErrorCodeMsg,
                                                          Optional ByVal fromReST As Boolean = False) _
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
            vpvListe = GETallVPvShort(vpid:="", err:=err,
                                      vpvid:="",
                                      status:="", refNext:=False,
                                      variantName:=noVariantName,
                                      storedAtorBefore:=storedAtOrBefore,
                                      fromReST:=True)

            For Each vpv As clsProjektWebShort In vpvListe
                Dim vpType As Integer = GETvpType(vpv.vpid, err)
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


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        retrieveProjectVariantNamesFromDB = result

    End Function

    ''' <summary>
    ''' holt Projekt-Namen über Angabe der Projekt-Nummer/Kundennummer beim Kunden; 
    ''' kann Null, ein oder mehrere Ergebnis-Einträge enthalten; Liste kommt sortiert nach Projekt-Namen zurück
    ''' </summary>
    ''' <param name="pNRatKD"></param>
    ''' <returns></returns>
    Public Function retrieveProjectNamesByPNRFromDB(ByVal pNRatKD As String, ByRef err As clsErrorCodeMsg) As Collection

        Dim result As New Collection
        Dim interimResult As New Collection

        Try

            Dim vpid As String = ""
            Dim anzLoop As Integer = 0

            'Dim allVP As New List(Of clsVP)
            While (result.Count <= 0 And anzLoop < 2)

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

                    VRScache.VPsId = GETallVP(aktVCid, err, ptPRPFType.all)

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
                                                 ByVal storedEarliest As DateTime, ByVal storedLatest As DateTime,
                                                 ByRef err As clsErrorCodeMsg) As clsProjektHistorie

        Dim result As New clsProjektHistorie
        storedLatest = storedLatest.ToUniversalTime()
        storedEarliest = storedEarliest.ToUniversalTime()

        Try

            'Dim zwischenResult As New SortedList(Of DateTime, clsProjektWebLong)
            Dim vpid As String = ""
            Dim vp As clsVP = GETvpid(projectname, err)

            ' VPID zu Projekt projectName holen vom WebServer/DB
            vpid = vp._id

            If vpid <> "" Then


                Dim allVPv As New List(Of clsProjektWebLong)
                ' erst alle mit dem angegebenen Varianten-NAmen holen 
                ' tk : ich habe das ausgeschrieben, das kann ich dann besser lesen - ausserdem passiert sonst leicht ein Fehler beim 'Abzählen' der optionalen Parameter
                ' allVPv = GETallVPvLong(vpid, err, , , , variantName)
                allVPv = GETallVPvLong(vpid:=vpid,
                                       err:=err,
                                       variantName:=variantName)

                ' einschränken auf alle versionen in dem angegebenen Zeitraum
                For Each vpv In allVPv
                    'If storedEarliest <= vpv.timestamp And vpv.timestamp <= storedLatest And vpv.variantName = variantName Then
                    If storedEarliest <= vpv.timestamp And vpv.timestamp <= storedLatest Then
                        'zwischenResult.Add(vpv.timestamp, vpv)
                        Dim hproj As New clsProjekt
                        vpv.copyto(hproj, vp)
                        result.Add(hproj.timeStamp, hproj)
                    End If
                Next

                ' jetzt alle Vorgaben holen, das sind die Versionen mit Varianten-NAme = "pfv" 
                'allVPv = GETallVPvLong(vpid, err, , , , ptVariantFixNames.pfv.ToString)
                allVPv = GETallVPvLong(vpid:=vpid,
                                       err:=err,
                                       variantName:=ptVariantFixNames.pfv.ToString)
                ' einschränken auf alle versionen in dem angegebenen Zeitraum

                For Each vpv In allVPv
                    ' die Vorgaben dürfen nicht an storedEarliest bzw storedlatest gebunden werden 
                    ' denn die können vor oder auch nach einem Planungs-Stand gespeichert worden sein 
                    'If storedEarliest <= vpv.timestamp And vpv.timestamp <= storedLatest Then

                    Dim hproj As New clsProjekt
                    vpv.copyto(hproj, vp)
                    result.AddPfv(hproj)

                    'End If
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
                                                     ByVal stored As DateTime, ByVal userName As String,
                                                 ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        If aktUser.email = userName Then

            stored = stored.ToUniversalTime.AddSeconds(1)

            Try
                Dim vpid As String = ""

                Dim vp As clsVP = GETvpid(projectname, err)
                ' VPID zu Projekt projectName holen vom WebServer/DB
                vpid = vp._id

                If vpid <> "" Then
                    ' gewünschte Variante vom Server anfordern
                    Dim allVPv As New List(Of clsProjektWebShort)
                    allVPv = GETallVPvShort(vpid:=vpid, err:=err,
                                            vpvid:="",
                                            status:="", refNext:=False,
                                            variantName:=variantName,
                                            storedAtorBefore:=stored,
                                            fromReST:=False)
                    If allVPv.Count >= 0 Then
                        If allVPv.Count = 1 Then
                            result = DELETEOneVPv(allVPv.Item(0)._id, err)
                        Else
                            For Each vpv As clsProjektWebShort In allVPv
                                If vpv.variantName = variantName Then
                                    result = result And DELETEOneVPv(vpv._id, err)
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
    Public Function retrieveFirstContractedPFromDB(ByVal projectname As String, ByVal variantname As String,
                                                   ByRef err As clsErrorCodeMsg) As clsProjekt

        Dim hproj As New clsProjekt

        Try
            Dim vpid As String = ""
            Dim vp As clsVP = GETvpid(projectname, err)

            If Not IsNothing(vp) Then

                ' VPID zu Projekt projectName holen vom WebServer/DB
                vpid = vp._id

                If vpid <> "" Then

                    Dim hresult As New List(Of clsProjektWebLong)

                    ' hresult kommt hier aufsteigend sortiert
                    hresult = GETallVPvLong(vpid:=vpid, err:=err, vpvid:="",
                                                status:="",
                                                refNext:=True,
                                                variantName:=variantname,
                                                storedAtorBefore:=Nothing)

                    ' das erste aus der Liste nehmen
                    If hresult.Count > 0 Then
                        hresult.Item(0).copyto(hproj, vp)
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
    ''' gibt den zum Zeitpunkt zuletzt beauftragten Stand zurück; 
    ''' </summary>
    ''' <param name="projectname"></param>
    ''' <param name="variantname"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <returns></returns>
    Public Function retrieveLastContractedPFromDB(ByVal projectname As String,
                                                  ByVal variantname As String,
                                                  ByVal storedAtOrBefore As DateTime,
                                                  ByRef err As clsErrorCodeMsg) As clsProjekt

        Dim hproj As New clsProjekt

        Try
            If (storedAtOrBefore = Date.MinValue) Then
                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime()
            Else
                storedAtOrBefore = storedAtOrBefore.ToUniversalTime()
            End If

            Dim vpid As String = ""
            Dim vp As clsVP = GETvpid(projectname, err)

            If Not IsNothing(vp) Then

                ' VPID zu Projekt projectName holen vom WebServer/DB
                vpid = vp._id

                If vpid <> "" Then

                    ' get specific VisboProjectVersion 
                    Dim hresult As New List(Of clsProjektWebLong)

                    ' hresult kommt hier aufsteigend sortiert
                    hresult = GETallVPvLong(vpid:=vpid, err:=err, vpvid:="",
                                            status:="",
                                            refNext:=False,
                                            variantName:=variantname,
                                            storedAtorBefore:=storedAtOrBefore)
                    ' das letzte aus der Liste nehmen
                    If hresult.Count > 0 Then
                        hresult.Item(hresult.Count - 1).copyto(hproj, vp)
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
    ''' <param name="type">  ptPRPFType.portfolio = 1
    '''                      ptPRPFType.project = 0
    '''                      ptPRPFType.projectTemplate = 2</param>
    ''' <returns>true -  es darf geändert werden
    '''          false - es darf nicht geändert werden</returns>
    Public Function checkChgPermission(ByVal pName As String, ByVal vName As String, ByVal userName As String, ByRef err As clsErrorCodeMsg, Optional type As Integer = ptPRPFType.project) As Boolean

        Dim result As Boolean = False

        Try

            Dim wpItem As clsWriteProtectionItem = getWriteProtection(pName, vName, err, type)

            If Not IsNothing(wpItem) Then
                If wpItem.isProtected Then
                    result = (wpItem.userName = aktUser.email)
                Else
                    result = True
                End If
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
    Public Function getWriteProtection(ByVal pName As String, ByVal vName As String,
                                       ByRef err As clsErrorCodeMsg,
                                       Optional type As ptPRPFType = ptPRPFType.project) As clsWriteProtectionItem
        Dim result As New clsWriteProtectionItem
        Try
            Dim vp As clsVP = GETvpid(pName, err, type)

            If Not IsNothing(vp) Then
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

            Else
                result = Nothing
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
    Public Function setWriteProtection(ByVal wpItem As clsWriteProtectionItem, ByRef err As clsErrorCodeMsg) As Boolean
        Dim result As Boolean = False

        Try
            If Not IsNothing(wpItem) Then
                Dim pname As String = Projekte.getPnameFromKey(wpItem.pvName)
                Dim vname As String = Projekte.getVariantnameFromKey(wpItem.pvName)

                Dim aktvp As clsVP = GETvpid(pname, err)

                If Not IsNothing(aktvp) Then

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
                            result = POSTVPLock(vpid, vname, err)
                        Else
                            result = DELETEVPLock(vpid, err, vname)
                        End If

                    Else

                        result = False

                    End If

                End If
            Else
                result = False
            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        setWriteProtection = result
    End Function

    Public Function retrieveProjectsOfOneConstellationFromDB(ByVal portfolioName As String,
                                                             ByRef err As clsErrorCodeMsg,
                                                             Optional ByVal storedAtOrBefore As Date = Nothing) As SortedList(Of String, clsProjekt)

        Dim result As New SortedList(Of String, clsProjekt)
        Dim intermediate As New List(Of clsProjektWebLong)
        Dim listOfPortfolios As New SortedList(Of Date, clsVPf)
        Dim vpid As String = ""
        Dim vptype As Module1.ptPRPFType = ptPRPFType.portfolio
        Dim vp As clsVP
        Dim vpfid As String = ""
        Dim hproj As New clsProjekt
        Try
            vp = GETvpid(portfolioName, err, vptype)
            vpid = vp._id
            listOfPortfolios = GETallVPf(vpid, storedAtOrBefore, err)
            vpfid = listOfPortfolios.Last.Value._id
            intermediate = GETallVPvOfOneVPf(aktVCid, vpfid, err, storedAtOrBefore, True)

            For Each webproj In intermediate

                hproj = New clsProjekt
                Dim thisVP As clsVP = GETvpid(webproj.name, err)
                webproj.copyto(hproj, thisVP)

                Dim a As Integer = hproj.dauerInDays
                Dim key As String = Projekte.calcProjektKey(hproj)
                If Not result.ContainsKey(key) Then
                    result.Add(key, hproj)
                End If
            Next

        Catch ex As Exception

        End Try


        retrieveProjectsOfOneConstellationFromDB = result
    End Function


    ''' <summary>
    '''  Alle Portfolios(Constellations) aus der Datenbank holen
    '''  Das Ergebnis dieser Funktion ist eine Liste (String, clsConstellation) 
    ''' </summary>
    ''' <returns></returns>
    Public Function retrieveConstellationsFromDB(ByRef err As clsErrorCodeMsg) As clsConstellations

        Dim result As New clsConstellations
        Try

            Dim intermediate As New SortedList(Of String, clsVP)
            Dim timestamp As Date = Date.Now.ToUniversalTime
            Dim c As New clsConstellation

            intermediate = GETallVP(aktVCid, err, ptPRPFType.portfolio)

            If err.errorCode = 200 Then

                For Each kvp As KeyValuePair(Of String, clsVP) In intermediate

                    If kvp.Value.vpType = ptPRPFType.portfolio Then

                        Dim vpid As String = kvp.Value._id
                        Dim portfolioVersions As SortedList(Of Date, clsVPf) = GETallVPf(vpid, timestamp, err)
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
                        ' kein Portfolio
                    End If

                Next
            End If


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
    Public Function storeConstellationToDB(ByVal c As clsConstellation,
                                           ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try
            Dim vpType As Integer = ptPRPFType.portfolio
            Dim cVPf As New clsVPf
            Dim cVP As New clsVP
            Dim newVP As New List(Of clsVP)
            Dim newVPf As New List(Of clsVPf)

            ' angepasst: 20180914: korrigieren, wenn ReST-Server geändert wurde
            '                       cVP = GETvpid(c.constellationName, vpType:=2)
            cVP = GETvpid(c.constellationName, err, ptPRPFType.portfolio)


            'cVPf = clsConst2clsVPf(c)

            If cVP._id = "" Then
                '' ur: war nur zu Testzwecken: 
                '' Call MsgBox("es ist noch kein VisboPortfolio angelegt")

                ' Portfolio-Name
                cVP.name = c.constellationName
                ' ur:14.12.2018: liste der User ist nicht mehr in den VPs enthalten
                '' berechtiger User
                'Dim user As New clsUser
                'user.email = aktUser.email
                'user.role = "Admin"
                'cVP.users.Add(user)
                ' VisboCenter - Id
                cVP.vcid = aktVCid
                ' VisboProject-Type - Portfolio
                cVP.vpType = ptPRPFType.portfolio

                ' Erzeugen des VisboPortfolios in der Collection visboproject im akt. VisboCenter
                newVP = POSTOneVP(cVP, err)
                If newVP.Count > 0 Then
                    cVP._id = newVP.Item(0)._id
                Else
                    Throw New ArgumentException("FEHLER beim erstellen des VisboPortfolioProject")
                End If

            End If

            cVPf = clsConst2clsVPf(c)

            If Not IsNothing(cVPf) Then
                cVPf.vpid = cVP._id

                ' timestamp setzen

                cVPf.timestamp = DateTimeToISODate(Date.UtcNow)


                If cVP._id <> "" Then

                    newVPf = POSTOneVPf(cVPf, err)

                    If newVPf.Count > 0 Then
                        result = True
                    End If

                End If
            Else

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
    Public Function removeConstellationFromDB(ByVal c As clsConstellation, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try
            Dim vpType As Integer = ptPRPFType.portfolio
            Dim cVPf As New clsVPf
            Dim cVP As New clsVP
            Dim newVP As New List(Of clsVP)
            Dim newVPf As New SortedList(Of Date, clsVPf)

            ' angepasst: 20180914: korrigieren, wenn ReST-Server geändert wurde
            'cVP = GETvpid(c.constellationName, vpType:=2)
            cVP = GETvpid(c.constellationName, err, ptPRPFType.portfolio)

            newVPf = GETallVPf(cVP._id, Date.Now.ToUniversalTime, err)

            'aktuell müssen zum löschen eines Portfolios alle PortfolioVersionen gelöscht werden
            If newVPf.Count > 0 Then

                If newVPf.Count = 1 Then
                    result = DELETEOneVPf(cVP._id, newVPf.ElementAt(0).Value._id, err)
                Else
                    Dim lv As Integer = 0
                    Dim ok As Boolean = True
                    result = ok
                    While result And (lv < newVPf.Count)
                        lv = lv + 1
                        ok = DELETEOneVPf(cVP._id, newVPf.ElementAt(lv - 1).Value._id, err)
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
                result = DELETEOneVP(cVP._id, err)
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
    Public Function retrieveWriteProtectionsFromDB(ByVal AlleProjekte As clsProjekteAlle, ByRef err As clsErrorCodeMsg) As SortedList(Of String, clsWriteProtectionItem)

        Dim result As New SortedList(Of String, clsWriteProtectionItem)

        Try
            For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

                Dim wpItem As New clsWriteProtectionItem
                wpItem.pvName = kvp.Key
                Dim pname As String = Projekte.getPnameFromKey(kvp.Key)
                Dim vname As String = Projekte.getVariantnameFromKey(kvp.Key)
                wpItem = getWriteProtection(pname, vname, err, ptPRPFType.project)

                If Not IsNothing(wpItem) Then
                    If Not result.ContainsKey(wpItem.pvName) Then
                        result.Add(wpItem.pvName, wpItem)
                    End If
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
    Public Function cancelWriteProtections(ByVal user As String, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Dim vplist As New SortedList(Of String, clsVP)

        Try
            ' alle vp des aktuellen Users und aktuellen vc holen
            If VRScache.VPsN.Count > 0 Then
                vplist = VRScache.VPsN
            Else
                vplist = GETallVP(aktVCid, err, ptPRPFType.all)
            End If

            For Each kvp As KeyValuePair(Of String, clsVP) In vplist

                If kvp.Value.lock.Count > 0 Then

                    ' holt zu der vpid die Varianten aus vpv Collection
                    Dim variantToProj As List(Of clsProjektWebShort) = GETallVPvShort(vpid:=kvp.Value._id, err:=err,
                                                                                      vpvid:="",
                                                                                      status:="", refNext:=False,
                                                                                      variantName:=noVariantName,
                                                                                      storedAtorBefore:=Date.Now.ToUniversalTime,
                                                                                      fromReST:=False)

                    ' Lock löschen für jede Variante des Projektes mit vpid
                    For Each vTp As clsProjektWebShort In variantToProj

                        result = result And DELETEVPLock(kvp.Value._id, err, vTp.variantName)

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
    Public Function retrieveRolesFromDB(ByVal storedAtOrBefore As DateTime, ByRef err As clsErrorCodeMsg) As clsRollen

        Dim result As New clsRollen()

        Try
            If storedAtOrBefore <= Date.MinValue Then
                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime()
            Else
                storedAtOrBefore = storedAtOrBefore.ToUniversalTime()
            End If

            Dim allRoles As New List(Of clsVCrole)

            ' Alle in der DB-vorhandenen Rollen mit timestamp <= refdate wäre wünschenswert
            allRoles = GETallVCrole(aktVCid, err)

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
    Public Function storeRoleDefinitionToDB(ByVal roleDef As clsRollenDefinition,
                                            ByVal insertNewDate As Boolean,
                                            ByVal ts As DateTime,
                                            ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try
            Dim timestamp As String = DateTimeToISODate(ts.ToUniversalTime())

            Dim role As New clsVCrole
            role.copyFrom(roleDef)
            role.timestamp = timestamp

            If insertNewDate Then
                result = POSTOneVCrole(aktVCid, role, err)
            Else
                If VRScache.VCrole.ContainsKey(role.name) Then
                    role._id = VRScache.VCrole(role.name)._id
                    result = PUTOneVCrole(aktVCid, role, err)
                End If

                If result = False Then ' Rolle ist noch nicht vorhanden im VisboCenter, also neu erzeugen
                    result = POSTOneVCrole(aktVCid, role, err)
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
    Public Function storeCostDefinitionToDB(ByVal costDef As clsKostenartDefinition,
                                            ByVal insertNewDate As Boolean,
                                            ByVal ts As DateTime,
                                            ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False

        Try
            Dim timestamp As String = DateTimeToISODate(ts.ToUniversalTime())

            Dim cost As New clsVCcost
            cost.copyFrom(costDef)
            cost.timestamp = timestamp

            If insertNewDate Then
                result = POSTOneVCcost(aktVCid, cost, err)
            Else

                If VRScache.VCcost.ContainsKey(cost.name) Then
                    cost._id = VRScache.VCcost(cost.name)._id
                    result = PUTOneVCcost(aktVCid, cost, err)
                End If

                If result = False Then  ' Kostenart ist noch nicht vorhanden im VisboCenter, also neu erzeugen
                    result = POSTOneVCcost(aktVCid, cost, err)
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
    Public Function retrieveCostsFromDB(ByVal storedAtOrBefore As DateTime, ByRef err As clsErrorCodeMsg) As clsKostenarten

        Dim result As New clsKostenarten()
        Try
            If storedAtOrBefore <= Date.MinValue Then
                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime()
            Else
                storedAtOrBefore = storedAtOrBefore.ToUniversalTime()
            End If

            Dim allCosts As New List(Of clsVCcost)
            ' Alle in der DB-vorhandenen Rollen mit timestamp <= refdate wäre wünschenswert
            allCosts = GETallVCcost(aktVCid, err)

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


    ''' <summary>
    ''' speichert eine VCSetting in der Datenbank; 
    ''' </summary>
    ''' <param name="listofSetting"></param>
    ''' <param name="type"></param>
    ''' <param name="ts"></param>
    ''' <param name="err"></param>
    ''' <returns></returns>
    Public Function storeVCsettingsToDB(ByVal listofSetting As Object,
                                        ByVal type As String,
                                        ByVal name As String,
                                        ByVal ts As DateTime,
                                        ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Dim setting As Object = Nothing
        Dim oldsetting As Object = Nothing
        Dim newsetting As Object = Nothing
        Dim settingID As String = ""
        Dim anzSetting As Integer = 0
        Dim timestamp As String = ""

        If ts > Date.MinValue Then
            ts = ts.ToUniversalTime.Date.AddDays(1)
            'ts = ts.ToUniversalTime
        End If

        Try

            Select Case type

                Case settingTypes(ptSettingTypes.customroles)
                    setting = New List(Of clsVCSettingCustomroles)
                    setting = GETOneVCsetting(aktVCid, type, name, Nothing, "", err, False)
                    anzSetting = CType(setting, List(Of clsVCSettingCustomroles)).Count
                    If anzSetting > 0 Then
                        settingID = CType(setting, List(Of clsVCSettingCustomroles)).ElementAt(0)._id
                    Else
                        settingID = ""
                    End If

                Case settingTypes(ptSettingTypes.customfields)
                    setting = New List(Of clsVCSettingCustomfields)
                    setting = GETOneVCsetting(aktVCid, type, name, Nothing, "", err, False)
                    anzSetting = CType(setting, List(Of clsVCSettingCustomfields)).Count
                    If anzSetting > 0 Then
                        settingID = CType(setting, List(Of clsVCSettingCustomfields)).ElementAt(0)._id
                    Else
                        settingID = ""
                    End If


                Case settingTypes(ptSettingTypes.organisation)
                    setting = New List(Of clsVCSettingOrganisation)
                    setting = GETOneVCsetting(aktVCid, type, name, ts, "", err, False)
                    anzSetting = CType(setting, List(Of clsVCSettingOrganisation)).Count
                    If anzSetting > 0 Then
                        oldsetting = CType(setting, List(Of clsVCSettingOrganisation)).ElementAt(0)
                        settingID = CType(setting, List(Of clsVCSettingOrganisation)).ElementAt(0)._id
                    Else
                        settingID = ""
                    End If

            End Select

            If ts > Date.MinValue Then
                timestamp = DateTimeToISODate(ts)
            Else
                timestamp = DateTimeToISODate(Date.Now.ToUniversalTime())
            End If

            Select Case type

                Case settingTypes(ptSettingTypes.customroles)

                    Dim listofCURsWeb As New clsCustomUserRolesWeb
                    listofCURsWeb.copyFrom(listofSetting)

                    ' der Unique-Key für customroles besteht aus: name, type

                    newsetting = New clsVCSettingCustomroles
                    CType(newsetting, clsVCSettingCustomroles).name = type         ' customroles '
                    CType(newsetting, clsVCSettingCustomroles).timestamp = timestamp
                    CType(newsetting, clsVCSettingCustomroles).userId = aktUser._id
                    CType(newsetting, clsVCSettingCustomroles).vcid = aktVCid
                    CType(newsetting, clsVCSettingCustomroles).type = type
                    CType(newsetting, clsVCSettingCustomroles).value = listofCURsWeb

                    If anzSetting = 1 Then
                        newsetting._id = settingID
                        ' Update der customroles - Setting
                        result = PUTOneVCsetting(aktVCid, settingTypes(ptSettingTypes.customroles), newsetting, err)
                    Else
                        ' Create der customroles - Setting
                        result = POSTOneVCsetting(aktVCid, settingTypes(ptSettingTypes.customroles), newsetting, err)
                    End If


                Case settingTypes(ptSettingTypes.customfields)

                    Dim listofCustomFields As New clsCustomFieldDefinitionsWeb
                    listofCustomFields.copyFrom(listofSetting)

                    ' der Unique-Key für customroles besteht aus: name, type

                    newsetting = New clsVCSettingCustomfields
                    CType(newsetting, clsVCSettingCustomfields).name = name         ' customfields-Date.now '
                    CType(newsetting, clsVCSettingCustomfields).timestamp = timestamp
                    CType(newsetting, clsVCSettingCustomfields).userId = aktUser._id
                    CType(newsetting, clsVCSettingCustomfields).vcid = aktVCid
                    CType(newsetting, clsVCSettingCustomfields).type = type
                    CType(newsetting, clsVCSettingCustomfields).value = listofCustomFields

                    If anzSetting = 1 Then
                        newsetting._id = settingID
                        ' Update der customroles - Setting
                        result = PUTOneVCsetting(aktVCid, settingTypes(ptSettingTypes.customfields), newsetting, err)
                    Else
                        ' Create der customroles - Setting
                        result = POSTOneVCsetting(aktVCid, settingTypes(ptSettingTypes.customfields), newsetting, err)
                    End If


                Case settingTypes(ptSettingTypes.organisation)

                    Dim listofOrgaWeb As New clsOrganisationWeb
                    listofOrgaWeb.copyFrom(listofSetting)

                    ' der Unique-Key für customroles besteht aus: name, type

                    newsetting = New clsVCSettingOrganisation
                    CType(newsetting, clsVCSettingOrganisation).name = name         ' Oranisation - ... '
                    Dim validFrom As String = DateTimeToISODate(listofOrgaWeb.validFrom)
                    CType(newsetting, clsVCSettingOrganisation).timestamp = validFrom
                    CType(newsetting, clsVCSettingOrganisation).userId = aktUser._id
                    CType(newsetting, clsVCSettingOrganisation).vcid = aktVCid
                    CType(newsetting, clsVCSettingOrganisation).type = type
                    CType(newsetting, clsVCSettingOrganisation).value = listofOrgaWeb

                    If anzSetting = 1 Then

                        ' Update der customroles - Setting
                        If CType(oldsetting, clsVCSettingOrganisation).value.validFrom = listofOrgaWeb.validFrom Then
                            newsetting._id = settingID
                            result = PUTOneVCsetting(aktVCid, settingTypes(ptSettingTypes.organisation), newsetting, err)
                        Else
                            ' Create der customroles - Setting
                            result = POSTOneVCsetting(aktVCid, settingTypes(ptSettingTypes.organisation), newsetting, err)
                        End If

                    Else
                        ' Create der customroles - Setting
                        result = POSTOneVCsetting(aktVCid, settingTypes(ptSettingTypes.organisation), newsetting, err)
                    End If
            End Select


            If err.errorCode <> 200 Then

                Select Case err.errorCode
                    Case 400
                    Case 401
                    Case 403
                    Case 409
                        ' PUTOneVCSetting erforderlich
                        Call MsgBox(err.errorMsg)
                    Case Else
                        Call MsgBox(err.errorMsg)
                End Select

            End If


        Catch ex As Exception
            'Throw New ArgumentException(ex.Message & err.errorMsg)
        End Try

        storeVCsettingsToDB = result
    End Function


    Public Function retrieveCustomUserRoles(ByRef err As clsErrorCodeMsg) As clsCustomUserRoles

        Dim result As New clsCustomUserRoles
        Dim setting As Object = Nothing
        Dim settingID As String = ""
        Dim anzSetting As Integer = 0
        Dim type As String = settingTypes(ptSettingTypes.customroles)
        Dim name As String = type
        Dim webCustomUserRoles As New clsCustomUserRolesWeb
        Try

            setting = New List(Of clsVCSettingCustomroles)
            setting = GETOneVCsetting(aktVCid, type, name, Nothing, "", err, False)
            If Not IsNothing(setting) Then

                anzSetting = CType(setting, List(Of clsVCSettingCustomroles)).Count

                If anzSetting > 0 Then

                    settingID = CType(setting, List(Of clsVCSettingCustomroles)).ElementAt(0)._id
                    webCustomUserRoles = CType(setting, List(Of clsVCSettingCustomroles)).ElementAt(0).value
                    webCustomUserRoles.copyTo(result)


                Else
                    result = New clsCustomUserRoles
                    'If err.errorCode = 403 Then
                    '    Call MsgBox(err.errorMsg)
                    'End If
                    settingID = ""
                End If

            Else

            End If


        Catch ex As Exception
            Call MsgBox(err.errorMsg)
        End Try
        retrieveCustomUserRoles = result
    End Function


    ''' <summary>
    ''' liest die komplette Organisation (Kosten und Rollen) aus den VCSettings
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="validfrom"></param>
    ''' <param name="err"></param>
    ''' <returns></returns>
    Public Function retrieveOrganisationFromDB(ByVal name As String,
                                         ByVal validfrom As Date,
                                         ByVal refnext As Boolean,
                                         ByRef err As clsErrorCodeMsg) As clsOrganisation

        Dim result As New clsOrganisation
        Dim setting As Object = Nothing
        Dim settingID As String = ""
        Dim anzSetting As Integer = 0
        Dim type As String = settingTypes(ptSettingTypes.organisation)

        validfrom = validfrom.ToUniversalTime

        Dim webOrganisation As New clsOrganisationWeb
        Try

            setting = New List(Of clsVCSettingOrganisation)
            setting = GETOneVCsetting(aktVCid, type, name, validfrom, "", err, refnext)

            If Not IsNothing(setting) Then

                anzSetting = CType(setting, List(Of clsVCSettingOrganisation)).Count

                If anzSetting > 0 Then
                    If anzSetting = 1 Then

                        settingID = CType(setting, List(Of clsVCSettingOrganisation)).ElementAt(0)._id
                        webOrganisation = CType(setting, List(Of clsVCSettingOrganisation)).ElementAt(0).value

                    Else
                        ' die Organisation suchen, die am nächsten an validFrom liegt
                        Dim latestOrga As New clsVCSettingOrganisation
                        Dim orgaSettingsListe As List(Of clsVCSettingOrganisation) = CType(setting, List(Of clsVCSettingOrganisation))

                        For Each orgaSetting As clsVCSettingOrganisation In orgaSettingsListe
                            If orgaSetting.timestamp > latestOrga.timestamp Then
                                If orgaSetting.timestamp <= validfrom Then
                                    latestOrga = orgaSetting
                                End If
                            End If
                        Next

                        webOrganisation = latestOrga.value

                    End If

                    webOrganisation.copyTo(result)

                    ' bestimmen der _topLevelNodeIDs
                    result.allRoles.buildTopNodes()

                Else
                    If err.errorCode = 403 Then
                        Call MsgBox(err.errorMsg)
                    End If
                    settingID = ""

                End If
            Else
                Call MsgBox(err.errorMsg)

            End If

        Catch ex As Exception

        End Try
        retrieveOrganisationFromDB = result
    End Function

    ''' <summary>
    ''' liest alle CustomFields aus VCSetting über ReST-Server
    ''' </summary>
    ''' <param name="err"></param>
    ''' <returns></returns>
    Public Function retrieveCustomFieldsFromDB(ByRef err As clsErrorCodeMsg) As clsCustomFieldDefinitions

        Dim result As New clsCustomFieldDefinitions
        Dim setting As Object = Nothing
        Dim settingID As String = ""
        Dim anzSetting As Integer = 0
        Dim type As String = settingTypes(ptSettingTypes.customfields)
        Dim name As String = type

        Dim webCustomFields As New clsCustomFieldDefinitionsWeb
        Try

            setting = New List(Of clsVCSettingCustomfields)
            setting = GETOneVCsetting(aktVCid, type, name, Nothing, "", err, False)

            If Not IsNothing(setting) Then

                anzSetting = CType(setting, List(Of clsVCSettingCustomfields)).Count

                If anzSetting > 0 Then

                    settingID = CType(setting, List(Of clsVCSettingCustomfields)).ElementAt(0)._id
                    webCustomFields = CType(setting, List(Of clsVCSettingCustomfields)).ElementAt(0).value
                    webCustomFields.copyTo(result)

                Else
                    If err.errorCode = 403 Then
                        Call MsgBox(err.errorMsg)
                    End If
                    settingID = ""
                End If

            End If


        Catch ex As Exception

        End Try
        retrieveCustomFieldsFromDB = result
    End Function

    Public Function retrieveUserIDFromName(ByVal username As String, ByRef err As clsErrorCodeMsg) As String

        Dim result As String = ""

        Try
            result = "not defined"
        Catch ex As Exception

        End Try
        retrieveUserIDFromName = result
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


        Dim response As HttpWebResponse = Nothing
        Dim hresp As HttpWebResponse = Nothing


        Dim proxyAuth As New frmProxyAuth   ' Formular zum erfragen der Proxy-Authentifizierung


        ''Dim registeredModules As IEnumerator = AuthenticationManager.RegisteredModules
        ''Call MsgBox("The following authentication modules are now registered with the system:")
        ''While registeredModules.MoveNext
        ''    Dim currentAuthenticationModule As IAuthenticationModule = registeredModules.Current
        ''    Call MsgBox("AuthenticateType: " & currentAuthenticationModule.AuthenticationType & vbLf &
        ''                "CanPreAuthenticate : " & currentAuthenticationModule.CanPreAuthenticate.ToString)
        ''End While

        Dim defaultProxy As IWebProxy = HttpWebRequest.DefaultWebProxy


        If awinSettings.visboDebug Then
            Dim proxyUri As Uri = defaultProxy.GetProxy(New Uri(awinSettings.databaseURL))
            Call MsgBox("ProxyURL zu " & awinSettings.databaseURL & " : " & proxyUri.ToString)
        End If


        Dim myProxy As New System.Net.WebProxy



        If awinSettings.proxyURL <> "" Then

            'prox.Address = New Uri("http://versicherung.proxy.allianz:8080")
            myProxy.Address = New Uri(awinSettings.proxyURL)

            Dim credentials As ICredentials = CredentialCache.DefaultNetworkCredentials

            '' Get the username And password from the credentials
            If Not IsNothing(netcred) Then
                If Not (netcred.UserName = "" Or netcred.Password = "") Then
                    myProxy.Credentials = netcred
                Else
                    Dim MyCreds As NetworkCredential = credentials.GetCredential(myProxy.Address, "Basic")
                    myProxy.Credentials = MyCreds
                End If

            Else
                Dim MyCreds As NetworkCredential = credentials.GetCredential(myProxy.Address, "Basic")
                myProxy.Credentials = MyCreds
            End If
        Else
            myProxy.Address = Nothing

        End If


        Dim credentialsErfragt As Boolean = False




        Try
            ' ur: 20190326: wird für tls1.2 benötigt - sicherer und ist in nginX definiert
            System.Net.ServicePointManager.Expect100Continue = True
            System.Net.ServicePointManager.SecurityProtocol =
            SecurityProtocolType.Tls Or
            SecurityProtocolType.Tls11 Or
            SecurityProtocolType.Tls12 Or
            SecurityProtocolType.Ssl3

            Dim request As HttpWebRequest = DirectCast(HttpWebRequest.Create(uri), HttpWebRequest)

            If IsNothing(myProxy.Address) Then
                request.Proxy = defaultProxy

            Else
                request.Proxy = myProxy
            End If


            request.UseDefaultCredentials = True
            'request.Credentials = CredentialCache.DefaultCredentials
            request.Credentials = CredentialCache.DefaultNetworkCredentials


            Dim cc As New CookieContainer
            request.CookieContainer = cc

            request.Method = method
            request.ContentType = visboContentType
            request.Headers.Add("access-key", token)
            request.UserAgent = visboUserAgent

            Dim toDo As Boolean = False
            Dim anzError As Integer = 0


            request.ContentLength = data.Length

            If request.ContentLength > 0 Then

                toDo = True

                While toDo And anzError < 3
                    Try
                        Using requestStream As Stream = request.GetRequestStream()

                            ' Send the data.
                            requestStream.Write(data, 0, data.Length)
                            requestStream.Close()
                            requestStream.Dispose()
                        End Using

                        If Not IsNothing(myProxy.Address) Then
                            ' ProxyURL merken
                            awinSettings.proxyURL = myProxy.Address.ToString
                        Else
                            ' Adresse von defaultProxy hier eintragen
                            awinSettings.proxyURL = defaultProxy.GetProxy(New Uri(awinSettings.databaseURL)).ToString

                            If awinSettings.proxyURL = awinSettings.databaseURL Then
                                ' es gibt keinen ProxyServer
                                awinSettings.proxyURL = ""
                            End If

                        End If


                        hresp = Nothing
                        toDo = False

                    Catch ex As WebException

                        anzError = anzError + 1


                        If ex.Status = WebExceptionStatus.ConnectFailure Then

                            request = DirectCast(HttpWebRequest.Create(uri), HttpWebRequest)
                            request.Method = method
                            request.ContentType = visboContentType
                            request.Headers.Add("access-key", token)
                            request.UserAgent = visboUserAgent


                            netcred = New NetworkCredential
                            Dim proxyName As String = ""

                            If awinSettings.proxyURL <> "" Then

                                'erneuter Versuch mit myProxy
                                proxyName = defaultProxy.GetProxy(New Uri(awinSettings.databaseURL)).ToString
                                If proxyName = awinSettings.databaseURL Then
                                    proxyName = ""
                                End If
                            Else
                                If Not IsNothing(myProxy.Address) Then
                                    proxyName = myProxy.Address.ToString
                                Else
                                    proxyName = ""
                                End If
                            End If

                            credentialsErfragt = askProxyAuthentication(proxyName, netcred.UserName, netcred.Password, netcred.Domain)

                            If proxyName <> "" And proxyName <> awinSettings.proxyURL Then
                                myProxy.Address = New Uri(proxyName)
                                request.Proxy = myProxy
                            End If

                            ' abgefragte Credentials beim Proxy eintragen
                            If Not IsNothing(request.Proxy) Then
                                request.Proxy.Credentials = netcred
                            End If

                        End If

                        If ex.Status = WebExceptionStatus.ProtocolError Then

                            hresp = ex.Response


                            If hresp.StatusCode = HttpStatusCode.ProxyAuthenticationRequired Then

                                request = DirectCast(HttpWebRequest.Create(uri), HttpWebRequest)
                                request.Method = method
                                request.ContentType = visboContentType
                                request.Headers.Add("access-key", token)
                                request.UserAgent = visboUserAgent

                                If credentialsErfragt And anzError = 2 Then
                                    Call MsgBox(hresp.Headers.ToString)
                                    Throw New ArgumentException("Fehler bei GetRequestStream:  " & vbCrLf & hresp.Headers.ToString & vbCrLf & ex.Message)
                                End If

                                Select Case anzError

                                    Case 1

                                        ' DefaultCredentials versuchen

                                        If myProxy.Address = Nothing Then
                                            request.Proxy = defaultProxy
                                        Else
                                            request.Proxy = myProxy
                                        End If

                                        request.UseDefaultCredentials = True
                                        request.Credentials = CredentialCache.DefaultCredentials
                                        'request.Credentials = CredentialCache.DefaultNetworkCredentials

                                    Case 2
                                        ' Abfragen der Proxy-Authentifizierung erforderlich

                                        netcred = New NetworkCredential
                                        Dim proxyName As String = ""

                                        If awinSettings.proxyURL <> "" Then
                                            proxyName = awinSettings.proxyURL
                                        Else
                                            If Not IsNothing(hresp) Then
                                                proxyName = hresp.ResponseUri.ToString
                                            End If

                                        End If

                                        credentialsErfragt = askProxyAuthentication(proxyName, netcred.UserName, netcred.Password, netcred.Domain)

                                        If proxyName <> "" And proxyName <> awinSettings.proxyURL Then
                                            myProxy.Address = New Uri(proxyName)
                                            request.Proxy = myProxy
                                        End If

                                        ' abgefragte Credentials beim Proxy eintragen
                                        If Not IsNothing(request.Proxy) Then
                                            request.Proxy.Credentials = netcred
                                        End If

                                        'credentialsErfragt = True 'zum Erkennen, ob Credentials für Proxy schon mal abgefragt wurden
                                        anzError = 1            ' wieder zurückgesetzt
                                End Select

                            Else
                                Throw New ArgumentException("Fehler bei GetRequestStream:  " & ex.Message)
                            End If
                        End If

                    End Try

                End While

            End If

            Dim fertig As Boolean = Not toDo

            If fertig Then

                If IsNothing(hresp) Then  ' hresp ist nur nicht nothing, wenn der request.GetRequestStream() schief ging

                    anzError = 0
                    toDo = True

                    While toDo And anzError < 3
                        Try
                            response = request.GetResponse()
                            toDo = False

                        Catch ex As WebException

                            anzError = anzError + 1

                            If ex.Status = WebExceptionStatus.ProtocolError Then

                                hresp = ex.Response
                                Select Case hresp.StatusCode

                                    Case HttpStatusCode.ProxyAuthenticationRequired

                                        request = DirectCast(HttpWebRequest.Create(uri), HttpWebRequest)
                                        request.Method = method
                                        request.ContentType = "application/json"
                                        request.Headers.Add("access-key", token)
                                        request.UserAgent = "VISBO Browser/x.x (" & My.Computer.Info.OSFullName & ":" & My.Computer.Info.OSPlatform & ":" _
                                                    & My.Computer.Info.OSVersion & ") Client:VISBO Projectboard/3.5 "

                                        Select Case anzError

                                            Case 1

                                                If myProxy.Address = Nothing Then
                                                    request.Proxy = defaultProxy
                                                Else
                                                    request.Proxy = myProxy
                                                End If

                                                request.UseDefaultCredentials = True
                                                request.Credentials = CredentialCache.DefaultCredentials


                                            Case 2
                                                ' Abfragen der Proxy-Authentifizierung erforderlich

                                                netcred = New NetworkCredential
                                                Dim proxyName As String = ""

                                                If awinSettings.proxyURL <> "" Then
                                                    proxyName = awinSettings.proxyURL
                                                End If

                                                credentialsErfragt = askProxyAuthentication(proxyName, netcred.UserName, netcred.Password, netcred.Domain)

                                                If proxyName <> "" And proxyName <> awinSettings.proxyURL Then
                                                    myProxy.Address = New Uri(proxyName)
                                                    request.Proxy = myProxy
                                                End If

                                                ' ur: für wingate-Proxy
                                                If Not IsNothing(request.Proxy) Then
                                                    request.Proxy.Credentials = netcred
                                                End If
                                        End Select
                                        'Case HttpStatusCode.BadRequest
                                        '    Exit While
                                        'Case HttpStatusCode.Unauthorized
                                        '    Exit While
                                        'Case HttpStatusCode.Forbidden
                                        '    Exit While
                                        'Case HttpStatusCode.NotFound
                                        '    Exit While
                                    Case Else
                                        response = hresp
                                        Exit While
                                End Select
                            End If

                        End Try

                    End While

                Else
                    response = hresp
                End If

            End If

            If anzError >= 3 Then
                response = hresp
            End If


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
    ''' <param name="err"></param>
    ''' <param name="vpType"></param>
    ''' <returns></returns>
    Private Function GETvpid(ByVal projectName As String,
                             ByRef err As clsErrorCodeMsg,
                             Optional ByVal vpType As Integer = ptPRPFType.project) As clsVP

        Dim vpid As String = ""
        Dim aktvp As New clsVP

        Try
            ' Alle VisboProjects über Server von WebServer/DB holen
            Dim anzLoop As Integer = 0
            'Dim allVP As New List(Of clsVP)
            While (vpid = "" And anzLoop < 3)

                If VRScache.VPsN.Count > 0 Then
                    ' Id zu angegebenen Projekt herausfinden
                    If VRScache.VPsN.ContainsKey(projectName) Then
                        Dim vp As clsVP = VRScache.VPsN.Item(projectName)
                        vpid = vp._id
                        aktvp = vp
                    Else
                        If anzLoop < 1 Then
                            VRScache.VPsN = GETallVP(aktVCid, err, vpType)
                        Else
                            VRScache.VPsN = GETallVP(aktVCid, err, ptPRPFType.all)

                        End If

                    End If
                Else
                    VRScache.VPsN = GETallVP(aktVCid, err, ptPRPFType.all)
                    If VRScache.VPsN.Count = 0 Then
                        anzLoop = 2
                    End If
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
        Dim err As New clsErrorCodeMsg

        Try
            ' Alle VisboProjects über Server von WebServer/DB holen
            Dim anzLoop As Integer = 0

            If vpid <> "" Then

                While (pName = "" And anzLoop < 2)

                    If VRScache.VPsId.Count > 0 Then

                        If VRScache.VPsId.ContainsKey(vpid) Then
                            ' pName zu angegebene vpid herausfinden
                            pName = VRScache.VPsId(vpid).name
                        Else

                            VRScache.VPsN = GETallVP(aktVCid, err, ptPRPFType.all)

                            If VRScache.VPsId.Count <= 0 Then
                                anzLoop = 2 ' while-Schleife beenden
                            Else
                                Try
                                    ' vermeiden, dass eine NullReference Exception geworfen wird ..
                                    If VRScache.VPsId.ContainsKey(vpid) Then
                                        pName = VRScache.VPsId(vpid).name
                                    Else
                                        pName = ""
                                    End If

                                Catch ex As Exception
                                    pName = ""
                                End Try
                            End If

                        End If
                    Else
                        VRScache.VPsN = GETallVP(aktVCid, err, ptPRPFType.all)

                        If VRScache.VPsId.Count <= 0 Then
                            anzLoop = 2 ' while-Schleife beenden
                        End If
                    End If

                    anzLoop = anzLoop + 1
                End While
            Else
                Throw New ArgumentException("Fehler in GETpName: vpid = "" übergeben")
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
    Private Function GETvpType(ByVal vpid As String, ByRef err As clsErrorCodeMsg) As Integer

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
                        VRScache.VPsN = GETallVP(aktVCid, err, ptPRPFType.all)

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
    Private Function GETallVP(ByVal vcid As String,
                              ByRef err As clsErrorCodeMsg,
                              Optional ByVal vptype As Integer = ptPRPFType.project) As SortedList(Of String, clsVP)

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

                            ' VPs nach Id sortiert gecacht
                            If Not VRScache.VPsId.ContainsKey(vp._id) Then
                                VRScache.VPsId.Add(vp._id, vp)
                            Else
                                VRScache.VPsId.Remove(vp._id)
                                VRScache.VPsId.Add(vp._id, vp)
                            End If

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

            err.errorCode = errcode
            err.errorMsg = "GETallVP" & " : " & errmsg & " : " & webVPantwort.message

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
                                    ByRef err As clsErrorCodeMsg,
                                   Optional vpvid As String = "",
                                   Optional status As String = "",
                                   Optional refNext As Boolean = False,
                                   Optional ByVal variantName As String = noVariantName,
                                   Optional ByVal storedAtorBefore As Date = Nothing,
                                   Optional ByVal fromReST As Boolean = False) As List(Of clsProjektWebShort)

        Dim nothingToDo As Boolean = True
        Dim result As New List(Of clsProjektWebShort)
        Dim errmsg As String = ""
        Dim errcode As Integer = 200

        If Not fromReST Then

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
        Else
            nothingToDo = False
        End If

        If nothingToDo Then

            If vpid <> "" And variantName <> noVariantName Then

                Dim variantlist As SortedList(Of Date, clsProjektWebShort) = VRScache.VPvs(vpid).Item(variantName).tsShort

                Dim found As Boolean = False
                Dim i As Integer = variantlist.Count - 1

                While Not found And i >= 0
                    Dim ts As Date = variantlist.ElementAt(i).Key
                    Dim shortproj As clsProjektWebShort = variantlist.ElementAt(i).Value

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
            Else
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

            End If

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
                            If variantName <> noVariantName Then
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
                                If variantName <> noVariantName Then
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
                                If variantName <> noVariantName Then
                                    serverUriString = serverUriString & "&variantName=" & variantName
                                End If

                            End If
                        End If

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

                err.errorCode = errcode
                err.errorMsg = "GETallVPvShort" & " : " & errmsg & " : " & webVPvAntwort.message

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
                                   ByRef err As clsErrorCodeMsg,
                                   Optional vpvid As String = "",
                                   Optional status As String = "",
                                   Optional refNext As Boolean = False,
                                   Optional ByVal variantName As String = noVariantName,
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

            If vpid <> "" And variantName <> noVariantName Then

                Dim variantlist As SortedList(Of Date, clsProjektWebLong) = VRScache.VPvs(vpid).Item(variantName).tsLong

                Dim found As Boolean = False
                Dim i As Integer = variantlist.Count - 1

                While Not found And i >= 0
                    Dim ts As Date = variantlist.ElementAt(i).Key
                    Dim longproj As clsProjektWebLong = variantlist.ElementAt(i).Value

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
            Else


                ' es existieren zu dieser vpid  und variantenName vpvs mit timestamps
                ' diese werden hier in die result-liste gebracht
                For Each kvp As KeyValuePair(Of String, SortedList(Of String, clsVarTs)) In VRScache.VPvs

                    Dim clsVarTs_vpid As String = kvp.Key

                    Dim clsVarTs_value As SortedList(Of String, clsVarTs) = VRScache.VPvs(clsVarTs_vpid)

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
                        If variantName <> noVariantName Then
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

            End If
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
                            If variantName <> noVariantName Then
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
                                If variantName <> noVariantName Then
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
                                If variantName <> noVariantName Then
                                    serverUriString = serverUriString & "&variantName=" & variantName
                                End If

                            End If
                        End If

                        ' es wird die Long-Version einer VisboProjectVersion angefordert
                        serverUriString = serverUriString & "&longList=1"

                    Else


                        ' Long-Version  angefordert
                        serverUriString = serverUriString & "&longList=1"


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

                err.errorCode = errcode
                err.errorMsg = "GETallVPvLong" & " : " & errmsg & " : " & webVPvAntwort.message

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
    Private Function GETallVPf(ByVal vpid As String, ByVal timestamp As Date, ByRef err As clsErrorCodeMsg) As SortedList(Of Date, clsVPf)

        Dim result As New SortedList(Of Date, clsVPf)          ' sortiert nach datum
        Dim secondResult As New SortedList(Of String, clsVPf)    ' sortiert nach vpid
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try
            Dim serverUriString As String
            Dim typeRequest As String = "/vp"

            ' URL zusammensetzen
            serverUriString = serverUriName & typeRequest & "/" & vpid & "/portfolio"

            If timestamp > Date.MinValue Then
                Dim refDate As String = DateTimeToISODate(timestamp)
                serverUriString = serverUriString & "?refDate=" & refDate
            End If

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

            err.errorCode = errcode
            err.errorMsg = "GETallVPf" & " : " & errmsg & " : " & webVPfantwort.message

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        GETallVPf = result

    End Function



    ''' <summary>
    ''' holt eine VisboPortfolio-Version und die zugehörigen Projekte/ProjektVersionen 
    ''' </summary>
    ''' <param name="vcid"></param>
    ''' <param name="vpfid"></param>
    ''' 
    ''' <param name="err"></param>
    ''' <returns>Liste von allen VPVs in diesem Portfolio</returns>
    Private Function GETallVPvOfOneVPf(ByVal vcid As String,
                                       ByVal vpfid As String,
                                       ByRef err As clsErrorCodeMsg,
                                       Optional ByVal storedAtorBefore As Date = Nothing,
                                       Optional ByVal longlist As Boolean = False) As List(Of clsProjektWebLong)

        Dim result As New List(Of clsProjektWebLong)
        Dim errmsg As String = ""
        Dim errcode As Integer


        Try
            Dim typeRequest As String = "/vpv"
            Dim serverUriString As String = serverUriName & typeRequest

            serverUriString = serverUriString & "?"
            serverUriString = serverUriString & "vcid=" & aktVCid

            If vpfid <> "" Then

                serverUriString = serverUriString & "&vpfid=" & vpfid

                If storedAtorBefore > Date.MinValue Then
                    Dim refDate As String = DateTimeToISODate(storedAtorBefore)
                    serverUriString = serverUriString & "&refDate=" & refDate
                End If

                If longlist Then
                    ' es wird die Long-Version einer VisboProjectVersion angefordert
                    serverUriString = serverUriString & "&longList=1"

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

                    '' cache soll nur befüllt werden, wenn nicht explizit mit VisboProjectVersion-Id aufgerufen
                    'If (vpfid = "") Then
                    '    ' nur dann soll der Cache gefüllt werden, damit auch wirklich alle aktuellen Timestamps enthalten sind
                    '    VRScache.createVPvLong(result, Date.Now.ToUniversalTime)
                    'End If

                Else

                    ' Fehlerbehandlung je nach errcode
                    Dim statError As Boolean = errorHandling_withBreak("GETOneVPfandAllVPvs", errcode, errmsg & " : " & webVPvAntwort.message)

                End If

                err.errorCode = errcode
                err.errorMsg = "GETOneVPfandAllVPvs" & " : " & errmsg & " : " & webVPvAntwort.message

            Else
                err.errorCode = 500
                err.errorCode = "Internal Error: vpfId nicht angegeben"
                '' Long-Version  angefordert
                'serverUriString = serverUriString & "&longList=1"

            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try


        GETallVPvOfOneVPf = result

    End Function



    ''' <summary>
    ''' löscht eine VisboProjectVersion
    ''' </summary>
    ''' <param name="vpvid"></param>
    ''' <returns>true:  löschen erfolgreich
    '''          false: löschen hat nicht funktioniert</returns>
    Private Function DELETEOneVPv(ByVal vpvid As String, ByRef err As clsErrorCodeMsg) As Boolean

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
                ' Cache aktualisieren
                VRScache.deleteVPv(vpvid)
            Else
                Dim statError As Boolean = errorHandling_withBreak("DELETEOneVPv", errcode, errmsg & " : " & webantwort.message)
            End If


            err.errorCode = errcode
            err.errorMsg = "DELETEOneVPv" & " : " & errmsg & " : " & webantwort.message

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
    Private Function DELETEOneVPf(ByVal vpid As String, ByVal vpfid As String, ByRef err As clsErrorCodeMsg) As Boolean

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


            err.errorCode = errcode
            err.errorMsg = "DELETEOneVPf" & " : " & errmsg & " : " & webantwort.message

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
    Private Function GETallVCrole(ByVal vcid As String, ByRef err As clsErrorCodeMsg) As List(Of clsVCrole)

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
            Dim webVCroleantwort As clsWebVCroles = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCroleantwort = JsonConvert.DeserializeObject(Of clsWebVCroles)(Antwort)
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

            err.errorCode = errcode
            err.errorMsg = "GETallVCrole" & " : " & errmsg & " : " & webVCroleantwort.message


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
    Private Function POSTOneVCrole(ByVal vcid As String, ByVal role As clsVCrole, ByRef err As clsErrorCodeMsg) As Boolean

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
            Dim webVCroleantwort As clsWebVCroles = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCroleantwort = JsonConvert.DeserializeObject(Of clsWebVCroles)(Antwort)
            End Using

            If errcode = 200 Then
                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTOneVCrole", errcode, errmsg & " : " & webVCroleantwort.message)
            End If


            err.errorCode = errcode
            err.errorMsg = "POSTOneVCrole" & " : " & errmsg & " : " & webVCroleantwort.message

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
    Private Function PUTOneVCrole(ByVal vcid As String, ByVal role As clsVCrole, ByRef err As clsErrorCodeMsg) As Boolean

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
            Dim webVCroleantwort As clsWebVCroles = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "PUT")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCroleantwort = JsonConvert.DeserializeObject(Of clsWebVCroles)(Antwort)
            End Using

            If errcode = 200 Then

                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("PUTOneVCrole", errcode, errmsg & " : " & webVCroleantwort.message)

            End If


            err.errorCode = errcode
            err.errorMsg = "PUTOneVCrole" & " : " & errmsg & " : " & webVCroleantwort.message

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
    Private Function GETallVCcost(ByVal vcid As String, ByRef err As clsErrorCodeMsg) As List(Of clsVCcost)

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
            Dim webVCcostantwort As clsWebVCcosts = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCcostantwort = JsonConvert.DeserializeObject(Of clsWebVCcosts)(Antwort)
            End Using

            If errcode = 200 Then

                result = webVCcostantwort.vccost
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("GETallVCcost", errcode, errmsg & " : " & webVCcostantwort.message)
            End If

            err.errorCode = errcode
            err.errorMsg = "GETallVCcost" & " : " & errmsg & " : " & webVCcostantwort.message


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
    Private Function POSTOneVCcost(ByVal vcid As String, ByVal cost As clsVCcost, ByRef err As clsErrorCodeMsg) As Boolean

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
            Dim webVCcostantwort As clsWebVCcosts = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCcostantwort = JsonConvert.DeserializeObject(Of clsWebVCcosts)(Antwort)
            End Using

            If errcode = 200 Then
                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTOneVCcost", errcode, errmsg & " : " & webVCcostantwort.message)
            End If

            err.errorCode = errcode
            err.errorMsg = "POSTOneVCcost" & " : " & errmsg & " : " & webVCcostantwort.message

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
    Private Function PUTOneVCcost(ByVal vcid As String, ByVal cost As clsVCcost, ByRef err As clsErrorCodeMsg) As Boolean

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
            Dim webVCcostantwort As clsWebVCcosts = Nothing
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "PUT")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                webVCcostantwort = JsonConvert.DeserializeObject(Of clsWebVCcosts)(Antwort)
            End Using

            If errcode = 200 Then

                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("PUTOneVCcost", errcode, errmsg & " : " & webVCcostantwort.message)

            End If

            err.errorCode = errcode
            err.errorMsg = "PUTOneVCcost" & " : " & errmsg & " : " & webVCcostantwort.message

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        PUTOneVCcost = result

    End Function


    ''' <summary>
    ''' liest ein Setting
    ''' </summary>
    ''' <param name="vcid"></param>
    ''' <param name="type"></param>
    ''' <param name="name"></param>
    ''' <param name="ts"></param>
    ''' <param name="userId"></param>
    ''' <param name="err"></param>
    ''' <returns></returns>
    Private Function GETOneVCsetting(ByVal vcid As String,
                                     ByVal type As String,
                                     ByVal name As String,
                                     ByVal ts As Date,
                                     ByVal userId As String,
                                     ByRef err As clsErrorCodeMsg,
                                     Optional ByVal refnext As Boolean = False) As Object

        Dim result As Object = Nothing
        Dim errmsg As String = ""
        Dim errcode As Integer
        Dim webVCsetting As Object = Nothing

        Try
            Dim timestamp As String = DateTimeToISODate(ts)

            Select Case type
                Case settingTypes(ptSettingTypes.customroles)
                    result = CType(result, clsVCSettingCustomroles)

                Case settingTypes(ptSettingTypes.customfields)
                    result = CType(result, clsVCSettingCustomfields)

                Case settingTypes(ptSettingTypes.organisation)
                    result = CType(result, clsVCSettingOrganisation)

                Case Else
                    Call MsgBox("settingType = " & type)
            End Select

            Dim serverUriString As String
            Dim typeRequest As String = "/vc"

            ' URL zusammensetzen
            If vcid = "" Then
                serverUriString = serverUriName & typeRequest
            Else
                serverUriString = serverUriName & typeRequest & "/" & vcid
            End If
            serverUriString = serverUriString & "/setting"

            If type <> "" Or name <> "" Or ts > Date.MinValue Then
                serverUriString = serverUriString & "?"


                If type <> "" Then
                    serverUriString = serverUriString & "type=" & type
                End If

                If name <> "" Then
                    serverUriString = serverUriString & "&name=" & name
                End If

                If name <> "" Or type <> "" Then
                    If ts > Date.MinValue Then
                        serverUriString = serverUriString & "&refDate=" & timestamp
                        If refnext Then
                            serverUriString = serverUriString & "&refNext=" & refnext.ToString
                        End If
                    Else
                        If refnext Then
                            serverUriString = serverUriString & "&refDate=" & timestamp
                            serverUriString = serverUriString & "&refNext=" & refnext.ToString
                        End If
                    End If
                End If

            End If


            Dim datastr As String = ""
            Dim encoding As New System.Text.UTF8Encoding()
            Dim data As Byte() = encoding.GetBytes(datastr)

            Dim serverUri As New Uri(serverUriString)

            Dim Antwort As String
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "GET")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                If errcode = 200 Then
                    Select Case type
                        Case settingTypes(ptSettingTypes.customroles)
                            webVCsetting = JsonConvert.DeserializeObject(Of clsWebVCSettingCustomroles)(Antwort)
                            result = CType(webVCsetting.vcsetting, List(Of clsVCSettingCustomroles))
                        Case settingTypes(ptSettingTypes.customfields)
                            webVCsetting = JsonConvert.DeserializeObject(Of clsWebVCSettingCustomfields)(Antwort)
                            result = CType(webVCsetting.vcsetting, List(Of clsVCSettingCustomfields))
                        Case settingTypes(ptSettingTypes.organisation)
                            webVCsetting = JsonConvert.DeserializeObject(Of clsWebVCSettingOrganisation)(Antwort)
                            result = CType(webVCsetting.vcsetting, List(Of clsVCSettingOrganisation))
                        Case Else
                            Call MsgBox("settingType = " & type)
                    End Select
                Else
                    webVCsetting = JsonConvert.DeserializeObject(Of clsWebOutput)(Antwort)
                End If

            End Using

            If errcode = 200 Then
                'nothing to do
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("GETOneVCsetting", errcode, errmsg & " : " & webVCsetting.message)
            End If


            err.errorCode = errcode
            err.errorMsg = "GETOneVCsetting" & " : " & errmsg & " : " & webVCsetting.message

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        GETOneVCsetting = result

    End Function

    ''' <summary>
    ''' erzeugt ein Setting
    ''' </summary>
    ''' <param name="vcid"></param>
    ''' <param name="setting"></param>
    ''' <returns></returns>
    Private Function POSTOneVCsetting(ByVal vcid As String, ByVal type As String, ByVal setting As Object, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer
        Dim webVCsetting As Object = Nothing


        Try

            Select Case type
                Case settingTypes(ptSettingTypes.customroles)
                    setting = CType(setting, clsVCSettingCustomroles)

                Case settingTypes(ptSettingTypes.customfields)
                    setting = CType(setting, clsVCSettingCustomfields)

                Case settingTypes(ptSettingTypes.organisation)
                    setting = CType(setting, clsVCSettingOrganisation)

                Case Else
                    Call MsgBox("Fehler: settingType = " & type & " íst nicht definiert")
            End Select

            Dim serverUriString As String
            Dim typeRequest As String = "/vc"

            ' URL zusammensetzen
            If vcid = "" Then
                serverUriString = serverUriName & typeRequest
            Else
                serverUriString = serverUriName & typeRequest & "/" & vcid
            End If
            serverUriString = serverUriString & "/setting"

            Dim serverUri As New Uri(serverUriString)
            Dim data As Byte() = serverInputDataJson(setting, "")



            Dim Antwort As String
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "POST")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                If errcode = 200 Then
                    Select Case type
                        Case settingTypes(ptSettingTypes.customroles)
                            webVCsetting = JsonConvert.DeserializeObject(Of clsWebVCSettingCustomroles)(Antwort)
                        Case settingTypes(ptSettingTypes.customfields)
                            webVCsetting = JsonConvert.DeserializeObject(Of clsWebVCSettingCustomfields)(Antwort)
                        Case settingTypes(ptSettingTypes.organisation)
                            webVCsetting = JsonConvert.DeserializeObject(Of clsWebVCSettingOrganisation)(Antwort)
                        Case Else
                            Call MsgBox("settingType = " & type)
                    End Select
                Else
                    webVCsetting = JsonConvert.DeserializeObject(Of clsWebOutput)(Antwort)
                End If

            End Using

            If errcode = 200 Then
                result = True
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTOneVCsetting", errcode, errmsg & " : " & webVCsetting.message)
            End If


            err.errorCode = errcode
            err.errorMsg = "POSTOneVCsetting" & " : " & errmsg & " : " & webVCsetting.message

        Catch ex As Exception
            'Throw New ArgumentException(ex.Message)
        End Try

        POSTOneVCsetting = result

    End Function


    ''' <summary>
    ''' update von Setting mit SettingID
    ''' </summary>
    ''' <param name="vcid"></param>
    ''' <param name="setting"></param>
    ''' <returns></returns>
    Private Function PUTOneVCsetting(ByVal vcid As String, ByVal type As String, ByRef setting As Object, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer
        Dim webVCsetting As Object = Nothing


        Try

            Select Case type
                Case settingTypes(ptSettingTypes.customroles)
                    setting = CType(setting, clsVCSettingCustomroles)

                Case settingTypes(ptSettingTypes.customfields)
                    setting = CType(setting, clsVCSettingCustomfields)

                Case settingTypes(ptSettingTypes.organisation)
                    setting = CType(setting, clsVCSettingOrganisation)

                Case Else
                    Call MsgBox("settingType = " & type)
            End Select

            Dim serverUriString As String
            Dim typeRequest As String = "/vc"

            ' URL zusammensetzen
            If vcid = "" Then
                serverUriString = serverUriName & typeRequest
            Else
                serverUriString = serverUriName & typeRequest & "/" & vcid
            End If
            serverUriString = serverUriString & "/setting/" & setting._id

            Dim serverUri As New Uri(serverUriString)
            Dim data As Byte() = serverInputDataJson(setting, "")



            Dim Antwort As String
            Using httpresp As HttpWebResponse = GetRestServerResponse(serverUri, data, "PUT")
                Antwort = ReadResponseContent(httpresp)
                errcode = CType(httpresp.StatusCode, Integer)
                errmsg = "( " & errcode.ToString & ") : " & httpresp.StatusDescription
                If errcode = 200 Then
                    Select Case type
                        Case settingTypes(ptSettingTypes.customroles)
                            webVCsetting = JsonConvert.DeserializeObject(Of clsWebVCSettingCustomroles)(Antwort)
                            setting = CType(webVCsetting.vcsetting, List(Of clsVCSettingCustomroles)).ElementAt(0)
                        Case settingTypes(ptSettingTypes.customfields)
                            webVCsetting = JsonConvert.DeserializeObject(Of clsWebVCSettingCustomfields)(Antwort)
                            setting = CType(webVCsetting.vcsetting, List(Of clsVCSettingCustomfields)).ElementAt(0)
                        Case settingTypes(ptSettingTypes.organisation)
                            webVCsetting = JsonConvert.DeserializeObject(Of clsWebVCSettingOrganisation)(Antwort)
                            setting = CType(webVCsetting.vcsetting, List(Of clsVCSettingOrganisation)).ElementAt(0)
                        Case Else
                            Call MsgBox("settingType = " & type)
                    End Select
                    result = True
                Else
                    webVCsetting = JsonConvert.DeserializeObject(Of clsWebOutput)(Antwort)
                    result = False
                End If

            End Using

            If errcode = 200 Then
                ' nothing to do
            Else
                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("PUTOneVCsetting", errcode, errmsg & " : " & webVCsetting.message)
            End If


            err.errorCode = errcode
            err.errorMsg = "PUTOneVCsetting" & " : " & errmsg & " : " & webVCsetting.message

        Catch ex As Exception
            'Throw New ArgumentException(ex.Message)
        End Try

        PUTOneVCsetting = result

    End Function


    ''' <summary>
    ''' ändert ein VisboProject
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird ein VisboProject geändert. user muss die Rechte haben, das checkt der Server</param>
    ''' <returns>Liste der VisboProjects</returns>
    Private Function PUTOneVP(ByVal vpid As String, ByVal vp As clsVP, ByRef err As clsErrorCodeMsg) As List(Of clsVP)

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
    ''' löscht das VP eines Projektes/variante
    ''' </summary>
    ''' <param name="vpid">vpid = "": es wird dass VisboProject vpid gelöscht. user muss die Rechte haben, das checkt der Server</param>
    ''' <returns>true: gelöscht
    '''          false: konnte nicht gelöscht werden</returns>
    Private Function DELETEOneVP(ByVal vpid As String, ByRef err As clsErrorCodeMsg) As Boolean

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

            err.errorCode = errcode
            err.errorMsg = "DELETEOneVP" & " : " & errmsg & " : " & webVP.message


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
    Private Function POSTVPLock(ByVal vpid As String,
                                ByVal variantName As String,
                                ByRef err As clsErrorCodeMsg) As Boolean


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

            err.errorCode = errcode
            err.errorMsg = "POSTVPLock" & " : " & errmsg & " : " & webVPLockantwort.message


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
    Private Function DELETEVPLock(ByVal vpid As String,
                                  ByRef err As clsErrorCodeMsg,
                                  Optional ByVal variantName As String = "") As Boolean

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

            err.errorCode = errcode
            err.errorMsg = "DELETEVPLock" & " : " & errmsg & " : " & webVPLockantwort.message

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
    Private Function POSTVPVariant(ByVal vpid As String, ByVal variantName As String,
                                   ByRef err As clsErrorCodeMsg) As Boolean


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

            err.errorCode = errcode
            err.errorMsg = "POSTVPVariant" & " : " & errmsg & " : " & webVPVar.message

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
    Private Function DELETEVPVariant(ByVal vpid As String,
                                     ByRef err As clsErrorCodeMsg,
                                     Optional ByVal varID As String = "") As Boolean

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

                err.errorCode = errcode
                err.errorMsg = "DELETEVPVariant" & " : " & errmsg & " : " & webVPVarAntwort.message

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
    Private Function POSTOneVP(ByVal vp As clsVP,
                               ByRef err As clsErrorCodeMsg) As List(Of clsVP)

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

            err.errorCode = errcode
            err.errorMsg = "POSTOneVP" & " : " & errmsg & " : " & webVPantwort.message

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        POSTOneVP = result

    End Function

    Private Function POSTOneVPv(ByVal vpid As String,
                                ByRef projekt As clsProjekt,
                                ByVal username As String, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        Dim errmsg As String = ""
        Dim errcode As Integer

        Try

            'Dim webVP As New clsWebVP
            'Dim vpErg As New List(Of clsVP)
            Dim data() As Byte

            Dim pname As String = projekt.name
            Dim vname As String = projekt.variantName

            ' jetzt muss noch VisboProjectVersion gespeichert werden
            Dim typeRequest As String = "/vpv"
            Dim serverUriString As String = serverUriName & typeRequest
            Dim serverUri As New Uri(serverUriString)


            If checkChgPermission(pname, vname, username, err) Then

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

                    ' vpv zu Cache hinzufügen
                    VRScache.createVPvLong(storeAntwort.vpv, Date.Now.ToUniversalTime)


                    If awinSettings.visboDebug Then

                        ' Rundum-Test
                        Dim newProjekt As New clsProjekt
                        Dim newWebProj As clsProjektWebLong = storeAntwort.vpv.ElementAt(0)

                        Dim vp As New clsVP
                        If VRScache.VPsId.ContainsKey(vpid) Then
                            vp = VRScache.VPsId(vpid)
                        End If

                        newWebProj.copyto(newProjekt, vp)
                        Dim korrekt As Boolean = newProjekt.isIdenticalTo(projekt)
                        If korrekt Then
                            Call MsgBox("Projekt nach POSTOneVPv gleich dem Ursprünglichen")
                        Else
                            Call MsgBox("FEHLER: Projekt nach POSTOneVPv nicht gleich dem Ursprünglichen")
                        End If

                    End If


                    ' updatedAt - Angabe in projekt speichern
                    If storeAntwort.vpv.Count >= 1 Then
                        projekt.updatedAt = storeAntwort.vpv.ElementAt(0).updatedAt
                    End If


                Else

                    ' Fehlerbehandlung je nach errcode
                    Dim statError As Boolean = errorHandling_withBreak("POSTOneVPv", errcode, errmsg & " : " & storeAntwort.message)

                End If
                err.errorCode = errcode
                err.errorMsg = "POSTOneVPv" & " : " & errmsg & " : " & storeAntwort.message
            End If

        Catch ex As Exception

        End Try
        POSTOneVPv = result
    End Function

    ''' <summary>
    ''' legt ein VisboPortfolio-Version an
    ''' </summary>
    ''' <param name="vpf">hier sind alle Daten des Projektes/Portfolios enthalten</param>
    ''' <returns>Liste mit dem angelegten VisboProject/VisboPortfolio inkl. kreierter _Id</returns>
    Private Function POSTOneVPf(ByVal vpf As clsVPf,
                                ByRef err As clsErrorCodeMsg) As List(Of clsVPf)

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

            err.errorCode = errcode
            err.errorMsg = "POSTOneVPf" & " : " & errmsg & " : " & webVPfantwort.message


        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        POSTOneVPf = result

    End Function

    Private Function POSTpwforgotten(ByVal ServerURL As String, ByVal databaseName As String,
                                     ByVal username As String, ByRef err As clsErrorCodeMsg) As Boolean

        Dim result As Boolean = False
        'Dim errmsg As String = ""
        'Dim errcode As Integer

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
                err.errorCode = CType(httpresp.StatusCode, Integer)
                err.errorMsg = "( " & err.errorCode.ToString & ") : " & httpresp.StatusDescription
                'webantwort = JsonConvert.DeserializeObject(Of clsWeboutput)(Antwort)
            End Using


            If err.errorCode = 200 Then

                result = True
            Else

                ' Fehlerbehandlung je nach errcode
                Dim statError As Boolean = errorHandling_withBreak("POSTpwforgotten", err.errorCode, err.errorMsg & " : " & webantwort.message)

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        POSTpwforgotten = result

    End Function
    '--------------------------------------------------------------------------------------------------------------------
    ' Allgemeine Funktionen und Prozeduren, die hierin benötigt werden
    '------------------------------------------------------------------------------------------------------------------

    ''' <summary>
    ''' Umwandlung einen Datum des Typs Date in einen ISO-Datums-String
    ''' </summary>
    ''' <param name="datumUhrzeit"></param>
    ''' <returns></returns>
    Public Function DateTimeToISODate(ByVal datumUhrzeit As Date) As String

        Dim ISODateandTime As String = Nothing
        Dim ISODate As String = ""
        Dim ISOTime As String = ""

        If datumUhrzeit >= Date.MinValue And datumUhrzeit <= Date.MaxValue Then
            ' DatumUhrzeit wird um 1 Sekunde erhöht, dass die 1000-stel keine Rolle spielen
            Dim hours As Integer = datumUhrzeit.Hour
            Dim minutes As Integer = datumUhrzeit.Minute
            Dim seconds As Integer = datumUhrzeit.Second
            Dim milliseconds As Integer = datumUhrzeit.Millisecond
            datumUhrzeit = datumUhrzeit.Date
            datumUhrzeit = datumUhrzeit.AddHours(hours).AddMinutes(minutes).AddSeconds(seconds).AddMilliseconds(0)
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

                    If hConstItem.projectName <> "" Then
                        Dim pvname As String = calcProjektKey(hConstItem.projectName, hConstItem.variantName)
                        If Not .Liste.ContainsKey(pvname) Then
                            result.Liste.Add(pvname, hConstItem)
                        End If
                    Else
                        ' der Fall kommt nur dann vor, wenn ein Portfolio mehrere Portfolios enthält, was nicht mehr sein darf
                        If awinSettings.visboDebug Then
                            Call MsgBox("Portfolio: " & vpf.name & vbCrLf & "ProjektID: " & hvpfItem.vpid)
                        End If
                    End If


                Next
                ' hier wird die Sortliste aufgebaut ... 
                .sortCriteria = vpf.sortType
                ' tk die Sort-Liste ist im Befehl vorher bereits aufgebaut 
                ' Dim hsortliste As SortedList(Of String, String) = .sortListe(vpf.sortType)
            End With

        Catch ex As Exception
            result = Nothing
        End Try

        clsVPf2clsConstellation = result

    End Function

    ''' <summary>
    ''' Kopieren des Portfolio c in das Portfolio des ReST-Servers vom Typ clsVPf
    ''' </summary>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Private Function clsConst2clsVPf(ByVal c As clsConstellation) As clsVPf

        Dim result As New clsVPf
        Dim err As New clsErrorCodeMsg
        Try
            Dim hvpid As String = ""
            Dim vpfItem As New clsVPfItem

            With result
                .name = c.constellationName
                ._id = ""

                ' angepasst: 20180914: ReST-Server muss auf ptPRPFType-Enumeration angepasst werden
                '.vpid = GETvpid(c.constellationName, vpType:=2)._id

                .vpid = GETvpid(c.constellationName, err, vpType:=ptPRPFType.portfolio)._id

                .timestamp = DateTimeToISODate(Date.Now)

                .sortType = c.sortCriteria
                ' .sortlist aufbauen aus c.sortlist
                For Each kvp As KeyValuePair(Of String, String) In c.sortListe(result.sortType)
                    hvpid = GETvpid(kvp.Value, err)._id

                    If hvpid = "" Then
                        result = Nothing   ' Signalisieren, dass ein Fehler aufgetaucht ist
                        Call MsgBox("neues Projekt '" & kvp.Value & "' bitte zuerst in DB speichern")
                        Throw New ArgumentException("neues Projekt '" & kvp.Value & "' bitte zuerst in DB speichern")
                    Else
                        If Not .sortList.Contains(hvpid) Then
                            .sortList.Add(hvpid)
                        End If
                    End If

                Next
                If Not IsNothing(result) Then
                    ' .allitems liste aufbauen aus c.allitems
                    For Each kvp As KeyValuePair(Of String, clsConstellationItem) In c.Liste
                        vpfItem = clsConstItem2clsVPfItem(kvp.Value)
                        If Not result.allItems.Contains(vpfItem) Then
                            result.allItems.Add(vpfItem)
                        End If
                    Next
                End If

            End With
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
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
                .projectTyp = vpfItem.reasonToExclude
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
        Dim err As New clsErrorCodeMsg
        Try
            With result

                result.name = cItem.projectName
                result.vpid = GETvpid(cItem.projectName, err)._id
                result._id = ""
                result.projectName = cItem.projectName
                result.variantName = cItem.variantName
                result.start = cItem.start
                result.show = cItem.show
                result.zeile = cItem.zeile
                result.reasonToExclude = cItem.projectTyp
                result.reasonToInclude = cItem.reasonToInclude

            End With

        Catch ex As Exception
            result = Nothing
        End Try

        clsConstItem2clsVPfItem = result

    End Function


    ''' <summary>
    ''' Leeren des VRSCache
    ''' </summary>
    ''' <returns></returns>
    Public Function clearVRSCache() As Boolean

        Dim result As Boolean = False
        Try
            ' Cache löschen, indem er neu aufgesetzt wird
            If Not IsNothing(VRScache) Then
                VRScache.Clear()
            Else
                VRScache = New clsCache
            End If
            result = True

        Catch ex As Exception
            result = False
        End Try


        clearVRSCache = result

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

                    If withBreak Then
                        Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)
                    End If


                Case 402        'Payment Required

                    If awinSettings.visboDebug Then
                        Call MsgBox("Fehler in " & restCall & " : " & webAntwortMsg)
                    End If
                    If withBreak Then
                        Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)
                    End If

                Case 403        ' Forbidden

                    If awinSettings.visboDebug Then
                        Call MsgBox("Fehler in " & restCall & " : " & webAntwortMsg)
                    End If
                    If withBreak Then
                        Throw New ArgumentException(errcode & ": Fehler in " & restCall & " : " & webAntwortMsg)
                    End If
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


End Class

