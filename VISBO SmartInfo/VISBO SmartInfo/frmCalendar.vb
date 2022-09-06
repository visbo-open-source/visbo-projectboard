Imports Microsoft.Office.Tools.Ribbon
Imports PPTNS = Microsoft.Office.Interop.PowerPoint
Imports System.Windows.Forms
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Public Class frmCalendar
    Private Sub frmCalendar_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Try
            ' Koordinaten merken
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try

        MyBase.Close()

    End Sub



    Private Sub frmCalendar_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call getFrmPosition(PTfrm.other, Top, Left)

        ' das Datum auf currentTimestamp setzen 
        DateTimePicker1.Value = currentTimestamp
        DateTimePicker2.Value = currentTimestamp


        'If englishLanguage Then
        '    ' Set the Format type and the CustomFormat string.
        '    DateTimePicker1.Format = DateTimePickerFormat.Custom
        '    DateTimePicker1.CustomFormat = "MMMM dd, yyyy - dddd"
        'Else
        '    DateTimePicker1.Format = DateTimePickerFormat.Long
        '    End If

    End Sub


    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click


        If DateTimePicker1.Value > Date.MinValue Then
            'DateTimePicker1.Value = DateTimePicker1.Value.Date.AddHours(23).AddMinutes(59)
            DateTimePicker1.Value = DateTimePicker1.Value.Date.AddHours(DateTimePicker2.Value.TimeOfDay.Hours).AddMinutes(DateTimePicker2.Value.TimeOfDay.Minutes).AddSeconds(DateTimePicker2.Value.TimeOfDay.Seconds)

        End If
        MyBase.Close()

    End Sub


End Class