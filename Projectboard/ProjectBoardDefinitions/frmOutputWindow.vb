Public Class frmOutputWindow

    Public textCollection As Collection
    Private Sub frmOutputWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Top = CInt(appInstance.Top + (appInstance.Height - Me.Height) / 2)
        Me.Left = CInt(appInstance.Left + (appInstance.Width - Me.Width) / 2)

        For i As Integer = 1 To textCollection.Count

            ' hier muss es jetzt zerhackt werden ... 
            Dim tmpstr() As String = CStr(textCollection.Item(i)).Split(New Char() {CChar(vbLf), CChar(vbCr)})

            For ii As Integer = 0 To tmpstr.Length - 1
                Me.ListBoxOutput.Items.Add(tmpstr(ii))
            Next


        Next

        Try
            Me.LinkLabelKontakt.Links.Add(0, 25, LinkLabelKontakt.Text)
        Catch ex As Exception

        End Try



    End Sub

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        textCollection = New Collection
        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub
    Private Sub linkLabelKontakt_LinkClicked(ByVal sender As Object,
                ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabelKontakt.LinkClicked

        ' Determine which link was clicked within the LinkLabel.
        Me.LinkLabelKontakt.Links(LinkLabelKontakt.Links.IndexOf(e.Link)).Visited = True

        ' Displays the appropriate link based on the value of the LinkData property of the Link object.
        Dim target As String = CType(e.Link.LinkData, String)

        ' If the value looks like a URL, navigate to it.
        ' Otherwise, display it in a message box.
        If (target IsNot Nothing) AndAlso (target.StartsWith("https")) Then
            System.Diagnostics.Process.Start(target)
        Else
            'Call MsgBox(("Item clicked: " + target))
        End If

    End Sub


End Class