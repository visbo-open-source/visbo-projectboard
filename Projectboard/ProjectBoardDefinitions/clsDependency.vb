''' <summary>
''' Klasse zur Beschreibung einer Abhängigkeit zwischen Projekten 
''' </summary>
''' <remarks></remarks>
Public Class clsDependency

    Public Property project As String
    Public Property dependentProject As String
    Public Property type As Integer
    Public Property degree As Integer
    Public Property description As String

    Public Sub New()
        _project = ""
        _dependentProject = ""
        _type = PTdpndncyType.inhalt
        _degree = PTdpndncy.schwach
        _description = ""
    End Sub

    Public Sub New(ByVal pname As String, ByVal dependentName As String, ByVal type As Integer, ByVal degree As Integer, ByVal description As String)
        _project = pname
        _dependentProject = dependentName
        _type = type
        _degree = degree
        _description = description
    End Sub


End Class
