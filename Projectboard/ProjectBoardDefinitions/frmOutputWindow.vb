Public Class frmOutputWindow

    Public textCollection As Collection
    Private Sub frmOutputWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For i As Integer = 1 To textCollection.Count

            ' hier muss es jetzt zerhackt werden ... 
            Dim tmpstr() As String = CStr(textCollection.Item(i)).Split(New Char() {CChar(vbLf), CChar(vbCr)})
            For ii As Integer = 0 To tmpstr.Length - 1
                ListBoxOutput.Items.Add(tmpstr(ii))
            Next

        Next

    End Sub

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        textCollection = New Collection
        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub
End Class