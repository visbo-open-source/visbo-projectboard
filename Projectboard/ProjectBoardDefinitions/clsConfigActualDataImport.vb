Public Class clsConfigActualDataImport
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
    Public Property sheet As clsRange
    Public Property sheetDescript As String
    Public Property column As clsRange
    Public Property columnDescript As String
    Public Property row As clsRange
    Public Property rowDescript As String
    Public Property objType As String
    Public Property content As String

    Public Sub New()
        Titel = ""
        Identifier = ""
        Inputfile = ""
        Typ = "Text"
        cellrange = False
        sheet = New clsRange
        sheetDescript = ""
        column = New clsRange
        columnDescript = "do not exist so far"
        row = New clsRange
        rowDescript = "do not exist so far"
        objType = ""
        content = """"

    End Sub
End Class
