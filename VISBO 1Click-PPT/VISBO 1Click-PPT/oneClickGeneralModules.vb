
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



Module oneClickGeneralModules

    Public pseudoappInstance As Microsoft.Office.Interop.Excel.Application

    Public Sub speichereProjektToDB(ByVal hproj As clsProjekt, _
                                    Optional ByVal messageZeigen As Boolean = False)

        Dim hprojVariante As String = ""

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
                            speichernInDBOk = storeSingleProjectToDB(hproj, identical)

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
                            speichernInDBOk = storeSingleProjectToDB(hproj, identical)
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

End Module
