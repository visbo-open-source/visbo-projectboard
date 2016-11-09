Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports MongoDbAccess



Module Module1

    Sub Main()
        ' (ByVal inputfile As String, ByVal username As String, ByVal password As String)
        Dim username As String = ""
        Dim password As String = ""

        Dim path As String = My.Computer.FileSystem.CurrentDirectory.ToString

        appInstance = New Microsoft.Office.Interop.Excel.Application
        Try
            If Not readawinSettings(path) Then

                awinSettings.databaseURL = My.Settings.mongoDBURL
                awinSettings.databaseName = My.Settings.mongoDBName
                awinSettings.globalPath = My.Settings.globalPath
                awinSettings.awinPath = My.Settings.awinPath

            End If

            Call awinsetTypen()

        Catch ex As Exception

            Call MsgBox(ex.Message)

        Finally
          
        End Try
        Call awinsetTypen()

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, username, password)

    End Sub


End Module
