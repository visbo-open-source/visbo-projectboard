Imports Microsoft.Office.Tools.Ribbon
Imports PPTNS = Microsoft.Office.Interop.PowerPoint
Imports System.Windows.Forms
Imports MongoDbAccess
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Public Class frmCalendar
    Private Sub frmCalendar_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        ' Koordinaten merken
        frmCoord(PTfrm.calendar, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.calendar, PTpinfo.left) = Me.Left
        frmCoord(PTfrm.calendar, PTpinfo.width) = Me.Width
        frmCoord(PTfrm.calendar, PTpinfo.height) = Me.Height

        If DateTimePicker1.Value > Date.MinValue Then
            DateTimePicker1.Value = DateTimePicker1.Value.Date.AddHours(23).AddMinutes(59)
        End If

        MyBase.Close()

    End Sub



    Private Sub frmCalendar_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If frmCoord(PTfrm.calendar, PTpinfo.top) > 0 Then
            Me.Top = frmCoord(PTfrm.calendar, PTpinfo.top)
            Me.Left = frmCoord(PTfrm.calendar, PTpinfo.left)
            Me.Width = frmCoord(PTfrm.calendar, PTpinfo.width)
            Me.Height = frmCoord(PTfrm.calendar, PTpinfo.height)
        End If


        'If englishLanguage Then
        '    ' Set the Format type and the CustomFormat string.
        '    DateTimePicker1.Format = DateTimePickerFormat.Custom
        '    DateTimePicker1.CustomFormat = "MMMM dd, yyyy - dddd"
        'Else
        '    DateTimePicker1.Format = DateTimePickerFormat.Long
        '    End If

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        ' MyBase.Close()

    End Sub

End Class