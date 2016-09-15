Public Class frmLoadConstellation

    Private formerselect As String
    Public retrieveFromDB As Boolean
    Public listOfTimeStamps As Collection
    Private Sub frmLoadConstellation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

            ListBox1.Items.Add(kvp.Key)

        Next
        formerselect = ""

        If Not retrieveFromDB Then
            dropBoxTimeStamps.Visible = False
            lblStandvom.Visible = False
        Else

            Try
                
                dropBoxTimeStamps.Items.Clear()

                For k As Integer = 1 To listOfTimeStamps.Count
                    Dim tmpDate As Date = CDate(listOfTimeStamps.Item(k))
                    dropBoxTimeStamps.Items.Add(tmpDate)
                Next

            Catch ex As Exception

            End Try

            ' jetzt ist dropBoxTimeStamps.selecteditem = Nothing ..
        End If

    End Sub

    Private Sub Abbrechen_Click(sender As Object, e As EventArgs) Handles Abbrechen.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        If ListBox1.SelectedItems.Count >= 1 Then
            DialogResult = System.Windows.Forms.DialogResult.OK
            MyBase.Close()
        Else
            Call MsgBox("bitte einen Eintrag selektieren")
        End If

    End Sub

    Private Sub addToSession_CheckedChanged(sender As Object, e As EventArgs) Handles addToSession.CheckedChanged


    End Sub

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        retrieveFromDB = False

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub
End Class