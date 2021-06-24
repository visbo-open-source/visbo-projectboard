Public Class clsConfigProjectsImport
    Public Class clsRange
        Public Property von As Integer
        Public Property bis As Integer
        Public Sub New()
            von = 0
            bis = 0
        End Sub
    End Class
    Public Property ProjectsFile As String
    Public Property hoursPerDay As Integer
    Public Property anzMonths As Integer


    Public Property Titel As String
    Public Property Identifier As String
    Public Property Inputfile As String
    Public Property Typ As String
    Public Property cellrange As Boolean
    Public Property sheet As Integer
    Public Property sheetDescript As String
    Public Property column As clsRange
    Public Property columnDescript As String
    Public Property row As clsRange
    Public Property rowDescript As String
    Public Property objType As String
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
        sheet = 0
        sheetDescript = ""
        column = New clsRange
        columnDescript = "do not exist so far"
        row = New clsRange
        rowDescript = "do not exist so far"
        objType = ""
        content = """"

    End Sub
End Class
