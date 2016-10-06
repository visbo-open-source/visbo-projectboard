Imports Microsoft.VisualStudio.Tools.Applications.Deployment
Imports Microsoft.VisualStudio.Tools.Applications

'Public Class FileCopyPDA
'    Implements IAddInPostDeploymentAction

'    Sub Execute(ByVal args As AddInPostDeploymentActionArgs) Implements IAddInPostDeploymentAction.Execute

'        Dim dataDirectory As String = "Projectboard.dll.config"
'        Dim file As String = "Projectboard.dll.config"
'        Dim sourcePath As String = args.AddInPath

'        Dim deploymentManifestUri As Uri = args.ManifestLocation
'        Dim destPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
'        Dim sourceFile As String = System.IO.Path.Combine(sourcePath, dataDirectory)
'        Dim destFile As String = System.IO.Path.Combine(destPath, file)


'        Select Case args.InstallationStatus
'            Case AddInInstallationStatus.InitialInstall, AddInInstallationStatus.Update
'                Try
'                    Call MsgBox("in Try")
'                    Call MsgBox("srcPath = " & sourcePath)
'                    Call MsgBox("sourceFile = " & sourceFile)
'                    Call MsgBox("destPath = " & destPath)
'                    Call MsgBox("destFile = " & destFile)

'                    System.IO.File.Copy(sourceFile, destFile)
'                    ServerDocument.RemoveCustomization(destFile)
'                    ServerDocument.AddCustomization(destFile, deploymentManifestUri)
'                Catch ex As Exception
'                    Call MsgBox("Catchmsg = " & ex.Message)
'                End Try

'                Exit Select
'            Case AddInInstallationStatus.Uninstall
'                If System.IO.File.Exists(destFile) Then
'                    System.IO.File.Delete(destFile)
'                End If
'                Exit Select
'        End Select
'    End Sub
'End Class
