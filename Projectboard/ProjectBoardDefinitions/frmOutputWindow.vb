Public Class frmOutputWindow

    Public textCollection As Collection
    Private Sub frmOutputWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For i As Integer = 1 To textCollection.Count

            ListBoxOutput.Items.Add(CStr(textCollection.Item(i)))

        Next


    End Sub

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        textCollection = New Collection
        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub
End Class