
Option Explicit On
'Option Strict On

Imports ProjectBoardDefinitions
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Forms
Imports System
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports Microsoft.VisualBasic
Imports ProjectBoardBasic
Imports System.Security.Principal
Imports DBAccLayer



Module oneClickGeneralModules

    Public pseudoappInstance As Microsoft.Office.Interop.Excel.Application

    Public Sub speichereProjektToDB(ByVal hproj As clsProjekt,
                                    Optional ByVal messageZeigen As Boolean = False)

        Dim hprojVariante As String = ""
        Dim outputCollection As New Collection
        Try
            ' LOGIN in DB machen
            If awinSettings.databaseURL <> "" And awinSettings.databaseName <> "" Then

                noDB = False

                If dbUsername = "" Or dbPasswort = "" Then

                    ' ur: 23.01.2015: Abfragen der Login-Informationen
                    loginErfolgreich = logInToMongoDB(True)

                    My.Settings.rememberUserPWD = awinSettings.rememberUserPwd
                    If awinSettings.rememberUserPwd Then
                        My.Settings.userNamePWD = awinSettings.userNamePWD
                    Else
                        My.Settings.userNamePWD = ""
                    End If


                    If Not loginErfolgreich Then
                        Call logfileSchreiben("LOGIN cancelled ...", "", -1)
                        Call MsgBox("LOGIN cancelled ...")
                    Else
                        Dim speichernInDBOk As Boolean = False
                        Dim identical As Boolean = False
                        Try
                            speichernInDBOk = storeSingleProjectToDB(hproj, outputCollection, identical:=identical)

                            If hproj.variantName <> "" Then
                                hprojVariante = "[" & hproj.variantName & "]"

                            End If

                            If speichernInDBOk Then
                                If messageZeigen Then
                                    If awinSettings.englishLanguage Then
                                        If Not identical Then
                                            Call MsgBox("Project '" & hproj.name & hprojVariante & "' stored" & vbLf & "version stored @ " & Date.Now.ToString)
                                        Else
                                            Call MsgBox("Project is identical to last database version" & vbLf & "no new version stored ")
                                        End If
                                    Else
                                        If Not identical Then
                                            Call MsgBox("Projekt '" & hproj.name & hprojVariante & "' gespeichert" & vbLf & "gespeicherte Version @ " & Date.Now.ToString)
                                        Else
                                            Call MsgBox("Projekt  ist identisch mit der aktuellen Version in der DB" & vbLf & "keine neue Version gespeichert")
                                        End If
                                    End If
                                End If
                            Else
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Error saving this project")
                                Else
                                    Call MsgBox("Fehler beim Speichern des aktuell geladenen Projektes")
                                End If

                            End If


                        Catch ex As Exception
                            If awinSettings.englishLanguage Then
                                Throw New ArgumentException("Error saving the project: " & hproj.name)
                            Else
                                Throw New ArgumentException("Fehler beim Speichern von Projekt: " & hproj.name)
                            End If

                        End Try

                    End If

                Else

                    If testLoginInfo_OK(dbUsername, dbPasswort) Then
                        Dim speichernInDBOk As Boolean
                        Dim identical As Boolean = False
                        Try
                            speichernInDBOk = storeSingleProjectToDB(hproj, outputCollection, identical)
                            If hproj.variantName <> "" Then
                                hprojVariante = "[" & hproj.variantName & "]"

                            End If
                            If speichernInDBOk Then

                                If messageZeigen Then
                                    If awinSettings.englishLanguage Then
                                        If Not identical Then
                                            Call MsgBox("Project '" & hproj.name & hprojVariante & "' stored" & vbLf & "version stored @ " & Date.Now.ToString)
                                        Else
                                            Call MsgBox("Project is identical to last database version" & vbLf & "no new version stored ")
                                        End If
                                    Else
                                        If Not identical Then
                                            Call MsgBox("Projekt '" & hproj.name & hprojVariante & "' gespeichert" & vbLf & "gespeicherte Version @ " & Date.Now.ToString)
                                        Else
                                            Call MsgBox("Projekt  ist identisch mit der aktuellen Version in der DB" & vbLf & "keine neue Version gespeichert")
                                        End If
                                    End If
                                End If

                            Else
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Error saving this project")
                                Else
                                    Call MsgBox("Fehler beim Speichern des aktuell geladenen Projektes")
                                End If
                            End If
                        Catch ex As Exception
                            If awinSettings.englishLanguage Then
                                Throw New ArgumentException("Error saving the project: " & hproj.name)
                            Else
                                Throw New ArgumentException("Fehler beim Speichern von Projekt: " & hproj.name)
                            End If
                        End Try
                    Else
                        If awinSettings.englishLanguage Then
                            Call MsgBox("LOGIN failure ...")
                        Else
                            Call MsgBox("LOGIN fehlerhaft ...")
                        End If

                    End If

                End If


            End If


        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' es wird der LoginProzess angestoßen. Bei erfolgreichem Login wird in den Settings verschlüsselt
    ''' userNamePWD gemerkt, sofern awinSettings.rememberUserPwd = true gesetzt ist.
    ''' Damit ist es möglich den nächsten Login zu automatisieren
    ''' </summary>
    ''' <param name="noDBAccess"></param>
    ''' <returns>true = erfolgreich</returns>
    Friend Function logInToMongoDB(ByVal noDBAccess As Boolean) As Boolean
        ' jetzt die Login Maske aufrufen, aber nur wenn nicht schon ein Login erfolgt ist .. ... 

        'awinSettings.visboServer = False ' Ohne Server

        ' bestimmt, ob in englisch oder auf deutsch ..
        Dim englishLanguage As Boolean = awinSettings.englishLanguage

        Dim msg As String = ""
        ''awinSettings.databaseURL = "http://visbo.myhome-server.de:3484"
        'awinSettings.databaseURL = "http://localhost:3484"
        'awinSettings.databaseName = "IT Projekte 2018"

        If awinSettings.databaseURL <> "" And awinSettings.databaseName <> "" Then

            ' jetzt prüfen , ob es bereits gespeicherte User-Credentials gibt 
            If IsNothing(awinSettings.userNamePWD) Then
                ' tk: 17.11.16: Einloggen in Datenbank 
                noDBAccess = Not loginProzedur()
                If Not noDBAccess Then
                    If awinSettings.rememberUserPwd Then
                        ' in diesem Fall das mySettings setzen 
                        Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)
                        awinSettings.userNamePWD = visboCrypto.verschluessleUserPwd(dbUsername, dbPasswort)
                    End If
                End If
            Else
                If awinSettings.userNamePWD = "" Then
                    ' tk: 17.11.16: Einloggen in Datenbank 
                    noDBAccess = Not loginProzedur()

                    If Not noDBAccess Then
                        If awinSettings.rememberUserPwd Then
                            ' in diesem Fall das mySettings setzen 
                            Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)
                            awinSettings.userNamePWD = visboCrypto.verschluessleUserPwd(dbUsername, dbPasswort)
                        End If
                    End If

                Else
                    ' die gespeicherten User-Credentials hernehmen, um sich einzuloggen 
                    ' ur: 19.06.2018
                    'noDBAccess = Not autoVisboLogin(awinSettings.userNamePWD)

                    ' wenn das jetzt nicht geklappt hat, soll wieder das login Fenster kommen ..
                    If noDBAccess Then
                        noDBAccess = Not loginProzedur()

                        If Not noDBAccess Then
                            If awinSettings.rememberUserPwd Then
                                ' in diesem Fall das mySettings setzen 
                                Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)
                                awinSettings.userNamePWD = visboCrypto.verschluessleUserPwd(dbUsername, dbPasswort)
                            End If
                        End If

                    End If

                End If

            End If

        End If

        If noDBAccess Then
            If englishLanguage Then
                msg = "no database access ... "
            Else
                msg = "kein Datenbank Zugriff ... "
            End If
            Call MsgBox(msg)
        Else
            ' hier müssen jetzt die Role- & Cost-Definitions gelesen werden 
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            'RoleDefinitions = request.retrieveRolesFromDB(currentTimestamp)
            'CostDefinitions = request.retrieveCostsFromDB(currentTimestamp)
            RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(Date.Now)
            CostDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCostsFromDB(Date.Now)
        End If

        logInToMongoDB = Not noDBAccess

    End Function



End Module
