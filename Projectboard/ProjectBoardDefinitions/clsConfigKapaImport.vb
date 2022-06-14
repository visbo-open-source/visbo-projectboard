Public Class clsConfigKapaImport
    Public Property capacityFile As String
    Public Property hoursPerDay As Integer
    Public Property anzMonths As Integer

    Public Property Titel As String
    Public Property Identifier As String
    Public Property Inputfile As String
    Public Property Typ As String
    Public Property cellrange As Boolean
    Public Property tabNr As Integer
    Public Property tabName As String
    Public Property column As Integer
    Public Property columnDescript As String
    Public Property row As Integer
    Public Property rowDescript As String
    Public Property regex As String
    Public Property content As String


    ''' <summary>
    ''' used when content carries offset Information for row, column
    ''' where to find the real values for the identifier field
    ''' need to be in form +(rowOffset, ColOffset)
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getRowColumnOffset() As Integer()
        Get
            Dim result As Integer()
            ReDim result(1)
            Try
                If content.StartsWith("+(") And
                    content.Contains(",") And
                    (content.EndsWith(")") Or content.EndsWith("):")) Then

                    Dim hstr() As String = content.Split(New Char() {CChar("("), CChar(","), CChar(")")})
                    result(0) = CInt(hstr(1))
                    result(1) = CInt(hstr(2))

                End If

            Catch ex As Exception

            End Try

            getRowColumnOffset = result
        End Get
    End Property


    Public Sub New()
        Titel = ""
        Identifier = ""
        Inputfile = ""
        Typ = "Text"
        cellrange = False
        tabNr = 1
        tabName = ""
        column = 0
        columnDescript = "do not exist so far"
        row = 0
        rowDescript = "do not exist so far"
        regex = ""
        content = """"

    End Sub

End Class
