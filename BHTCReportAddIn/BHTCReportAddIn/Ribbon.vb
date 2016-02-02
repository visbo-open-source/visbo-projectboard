'TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

'1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon()
'End Function

'2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
'   zu behandeln, zum Beispiel das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem
'   Menüband-Designer exportiert haben, verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und
'   ändern Sie den Code für die Verwendung mit dem Programmiermodell für die Menübanderweiterung (RibbonX).

'3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.

'Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.

<Runtime.InteropServices.ComVisible(True)> _
    Public Class Ribbon
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("BHTCReportAddIn.Ribbon.xml")
    End Function

#Region "Menübandrückrufe"
    'Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter "http://go.microsoft.com/fwlink/?LinkID=271226".
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub



#End Region

#Region "Hilfsprogramme"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
