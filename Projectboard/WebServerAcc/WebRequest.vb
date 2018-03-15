
Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports System.Windows
Imports System
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports System.Drawing
Imports System.Globalization
Imports System.Web
Imports System.ServiceModel.Web
Imports Microsoft.VisualBasic
Imports System.Security.Principal
Imports System.Net
Imports System.Text
Public Module WebRequest

    Public token As String = ""
    ''' <summary>
    ''' Sendet einen Request an den Server. Außerdem wird hier auch die Antwort empfangen und an die aufrufenden Routine zurückgegeben
    ''' </summary>
    ''' <param name="uri"></param>
    ''' <param name="data"></param>
    ''' <param name="callback"></param>
    Function GetPOSTResponse(uri As Uri, data As Byte(), callback As Action(Of HttpWebResponse)) As HttpWebResponse

        Dim response As HttpWebResponse = Nothing
        Try
            Dim request As HttpWebRequest = DirectCast(HttpWebRequest.Create(uri), HttpWebRequest)

            request.Method = "POST"
            request.ContentType = "application/json"
            request.Headers.Add("access-key", token)
            request.UserAgent = "VISBO Browser/x.x (" & My.Computer.Info.OSFullName & ":" & My.Computer.Info.OSPlatform & ":" & My.Computer.Info.OSVersion & ") Client:VISBO Projectboard/3.5 "

            request.ContentLength = data.Length
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


            Try
                response = request.GetResponse()

            Catch ex As WebException
                response = ex.Response
            End Try

            ''''Try

            ''''    request.BeginGetResponse(
            ''''    Function(x)
            ''''        Try
            ''''            response = DirectCast(request.EndGetResponse(x), HttpWebResponse)
            ''''            Return response
            ''''        Catch ex As WebException
            ''''            Using Exresponse As WebResponse = ex.Response
            ''''                Dim httpResponse As HttpWebResponse = DirectCast(Exresponse, HttpWebResponse)
            ''''                System.Diagnostics.Debug.WriteLine("Error code: {0}", httpResponse.StatusCode)
            ''''                Using str As Stream = Exresponse.GetResponseStream()
            ''''                    Dim sr = New StreamReader(str)
            ''''                    Dim text As String = sr.ReadToEnd()
            ''''                    System.Diagnostics.Debug.WriteLine(text)
            ''''                End Using
            ''''            End Using
            ''''            Return 0
            ''''        Catch ex As Exception
            ''''            System.Diagnostics.Debug.WriteLine("Message: " & ex.Message)
            ''''            Return 0
            ''''        End Try

            ''''    End Function, request)

            ''''Catch ex As Exception
            ''''    Call MsgBox("Fehler bei BeginGetResponse:  " & ex.Message)
            ''''    Return Nothing
            ''''End Try

        Catch ex1 As Exception
            Call MsgBox(ex1.Message)
            Throw
        End Try

        Return response

    End Function

    ''' <summary>
    ''' Sendet einen Request an den Server. Außerdem wird hier auch die Antwort empfangen und an die aufrufenden Routine zurückgegeben
    ''' </summary>
    ''' <param name="uri"></param>
    ''' <param name="data"></param>
    ''' <param name="callback"></param>
    Function GetGETResponse(uri As Uri, data As String, callback As Action(Of HttpWebResponse)) As HttpWebResponse

        Dim response As HttpWebResponse = Nothing
        Try

            Dim request As HttpWebRequest = DirectCast(HttpWebRequest.Create(uri), HttpWebRequest)

            request.Method = "GET"
            request.Headers.Add("access-key", token)
            request.Accept = "application/json"
            request.UserAgent = "VISBO Browser/x.x (" & My.Computer.Info.OSFullName & ":" & My.Computer.Info.OSPlatform & ":" & My.Computer.Info.OSVersion & ":" & myWindowsName & ") Client:VISBO Projectboard/3.5 "



            'request.ContentType = "application/json"
            'request.Headers.Add("access-key", token)
            'request.PreAuthenticate = True
            'request.Headers.Add("Cache-Control", "no-cache")

            ' nur notwendig, wenn ein Body mit übergeben wird

            'Dim encoding As New System.Text.UTF8Encoding()
            'Dim bytes As Byte() = encoding.GetBytes(data)

            'request.ContentLength = bytes.Length
            'Try
            '    Using requestStream As Stream = request.GetRequestStream()
            '        ' Send the data.
            '        requestStream.Write(bytes, 0, bytes.Length)
            '        requestStream.Close()
            '        requestStream.Dispose()
            '    End Using
            'Catch ex As Exception
            '    Call MsgBox("Fehler bei GetRequestStream:   " & ex.Message)
            '    Throw New ArgumentException("Fehler bei GetRequestStream:  " & ex.Message)
            'End Try


            Try
                response = request.GetResponse()

            Catch ex As WebException
                response = ex.Response
            End Try

            ''''Try

            ''''    request.BeginGetResponse(
            ''''    Function(gx)
            ''''        Try
            ''''            response = DirectCast(request.EndGetResponse(gx), HttpWebResponse)
            ''''            Return response
            ''''        Catch ex As WebException
            ''''            Using Exresponse As WebResponse = ex.Response
            ''''                Dim httpResponse As HttpWebResponse = DirectCast(Exresponse, HttpWebResponse)
            ''''                System.Diagnostics.Debug.WriteLine("Error code: {0}", httpResponse.StatusCode)
            ''''                Using str As Stream = Exresponse.GetResponseStream()
            ''''                    Dim sr = New StreamReader(str)
            ''''                    Dim text As String = sr.ReadToEnd()
            ''''                    System.Diagnostics.Debug.WriteLine(text)
            ''''                End Using
            ''''            End Using
            ''''            Return 0
            ''''        Catch ex As Exception
            ''''            System.Diagnostics.Debug.WriteLine("Message: " & ex.Message)
            ''''            Return 0
            ''''        End Try

            ''''    End Function, request)

            ''''Catch ex As Exception
            ''''    Call MsgBox("Fehler bei BeginGetResponse:  " & ex.Message)
            ''''    Return Nothing
            ''''End Try

        Catch ex1 As Exception
            Call MsgBox(ex1.Message)
            Throw
        End Try

        If IsNothing(response) Then
            Throw New HttpException(HttpStatusCode.NotFound, "The requested url could not be found.")
        End If
        Return response

    End Function

    Function ReadResponseContent(ByRef resp As HttpWebResponse) As String
        If IsNothing(resp) Then
            Throw New ArgumentNullException("resp")
        Else
            Using sr As New StreamReader(resp.GetResponseStream)
                Return sr.ReadToEnd()
            End Using
        End If
    End Function





    ''' <summary>
    ''' Es wird die Antwort des WebServers auf den Request vom Typ type in die jeweils entsprechende Klasse zerlegt (mit JsonSerializer
    ''' Ergebnis: Object in passender Struktur 
    ''' </summary>
    ''' <param name="resp"></param>
    ''' <param name="type"></param>
    ''' <returns>Object</returns>
    Function ReadGETResponseContentJson(ByRef resp As HttpWebResponse, ByVal type As String) As Object


        ReadGETResponseContentJson = Nothing

        If IsNothing(resp) Then
            Throw New ArgumentNullException("resp")
        Else
            Select Case type

                Case "/token/user/signin"

                Case "/token/user/login"

                    Dim tokenUserLogin As clsTokenUserLogin
                    Dim serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(GetType(clsTokenUserLogin))
                    Try
                        tokenUserLogin = serializer.ReadObject(resp.GetResponseStream)
                        ReadGETResponseContentJson = tokenUserLogin
                    Catch ex As Exception
                        Call MsgBox("Fehler in ReadResponseContent /token/user/login: " & ex.Message)
                    End Try

                Case "/user/changepw"

                Case "/user/forgotpw"

                Case "/vc"

                    Dim allVC As clsWebAllVC
                    Dim serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(GetType(clsWebAllVC))
                    Try
                        allVC = serializer.ReadObject(resp.GetResponseStream)
                        ReadGETResponseContentJson = allVC
                    Catch ex As Exception
                        Call MsgBox("Fehler in ReadGETResponseContent /vc: " & ex.Message)
                    End Try


            End Select


        End If
    End Function




    ''' <summary>
    ''' Es wird die Antwort des WebServers auf den Request vom Typ type in die jeweils entsprechende Klasse zerlegt (mit JsonSerializer
    ''' Ergebnis: Object in passender Struktur 
    ''' </summary>
    ''' <param name="resp"></param>
    ''' <param name="type"></param>
    ''' <returns>Object</returns>
    Function ReadPOSTResponseContentJson(ByRef resp As HttpWebResponse, ByVal type As String) As Object


        ReadPOSTResponseContentJson = Nothing

        If IsNothing(resp) Then
            Throw New ArgumentNullException("resp")
        Else
            Select Case type

                Case "/token/user/signin"

                Case "/token/user/login"

                    Dim tokenUserLogin As clsTokenUserLogin
                    Dim serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(GetType(clsTokenUserLogin))
                    Try
                        tokenUserLogin = serializer.ReadObject(resp.GetResponseStream)
                        ReadPOSTResponseContentJson = tokenUserLogin
                    Catch ex As Exception
                        Call MsgBox("Fehler in ReadPOSTResponseContent /token/user/login: " & ex.Message)
                    End Try

                Case "/user/changepw"

                Case "/user/forgotpw"

                Case "/vc"

                    Dim oneVC As clsWebOneVC
                    Dim serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(GetType(clsWebOneVC))
                    Try
                        oneVC = serializer.ReadObject(resp.GetResponseStream)
                        ReadPOSTResponseContentJson = oneVC
                    Catch ex As Exception
                        Call MsgBox("Fehler in ReadPOSTResponseContent /vc: " & ex.Message)
                    End Try


            End Select


        End If
    End Function





    ''' <summary>
    ''' diese Funktion konvertiert die Struktur, die für diesen Server-Request benötigt wird (type) in ein ByteArray im Json-Format
    ''' </summary>
    ''' <param name="dataClass"></param>
    ''' <param name="type"></param>
    ''' <returns>Object</returns>
    Function serverInputDataJson(ByVal dataClass As Object, ByVal type As String) As Byte()


        serverInputDataJson = Nothing

        If IsNothing(dataClass) Then
            Throw New ArgumentNullException("dataClass")
        Else
            Select Case type

                Case "/token/user/login"

                    Dim tokenUserLogin As clsTokenUserLogin
                    Dim serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(GetType(clsTokenUserLogin))
                    Try

                    Catch ex As Exception
                        Call MsgBox("Fehler in WriteJson /token/user/login: " & ex.Message)
                    End Try

                Case "/vc"

                    Dim teststring As String = ""
                    Dim bytes() As Byte = Nothing
                    Dim serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(GetType(clsVC))
                    Dim bufferlge As Int32 = 256
                    Dim ms As New MemoryStream(bufferlge)

                    Try
                        serializer.WriteObject(ms, dataClass)
                        ReDim bytes(ms.Length)
                        bytes = ms.GetBuffer()

                        Dim encoding As New System.Text.UTF8Encoding()
                        Dim hstr As String = encoding.GetString(bytes)
                        Call MsgBox(hstr)

                        ms.Close()
                    Catch ex As Exception
                        Call MsgBox("Fehler in writeJson /vc: " & ex.Message)
                    End Try

                    serverInputDataJson = bytes

                Case Else
                    Call MsgBox("Es ist wohl ein Fehler aufgetreten")
            End Select


        End If
    End Function



    ''' <summary>
    ''' Test Sub: Errechnete Struktur einer WebResponse in File namefile exportieren in Json
    ''' </summary>
    ''' <param name="clsJson"></param>
    ''' <param name="namefile"></param>
    Sub JsonExport(ByVal clsJson As clsWebAllVC, ByVal namefile As String)


        Dim jsonfilename As String = awinPath & namefile

        Try
            Dim serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(GetType(clsWebAllVC))

            Dim file As New FileStream(jsonfilename, FileMode.Create)
            serializer.WriteObject(file, clsJson)
            file.Close()

        Catch ex As Exception
            Call MsgBox("Beim Schreiben der Json-Datei '" & jsonfilename & "' ist ein Fehler aufgetreten !")
        End Try

    End Sub



    ''' <summary>
    ''' Lesen eines Json-Files von bestimmter Struktur
    ''' </summary>
    ''' <param name="namefile"></param>
    ''' <returns></returns>
    Function JsonImport(ByVal namefile As String) As clsWebAllVC
        Dim resp As HttpWebResponse = Nothing

        Dim tokenLogin As New clsWebAllVC

        Dim serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(GetType(clsWebAllVC))
        Dim jsonfilename As String = awinPath & namefile
        Try

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.

            Dim file As New FileStream(jsonfilename, FileMode.Open)
            tokenLogin = serializer.ReadObject(file)

            'Dim file As New StreamReader(resp.GetResponseStream)
            'tokenLogin = serializer.ReadObject(resp.GetResponseStream)


            JsonImport = tokenLogin

        Catch ex As Exception
            'Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            Throw New ArgumentException("Beim Lesen der Json-Datei '" & jsonfilename & "' ist ein Fehler aufgetreten !")
            JsonImport = Nothing
        End Try

    End Function

End Module
