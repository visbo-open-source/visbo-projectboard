Public Class clsConfigLine

    Public Property Titel As String
    Public Property Identifier As String
    Public Property InputFile As String
    Public Property Typ As String
    Public Property Datenbereich As String
    Public Property TabellenNummer As String
    Public Property TabellenName As String
    Public Property SpaltenNummer As String
    Public Property Spaltenüberschrift As String
    Public Property ZeilenNummer As String
    Public Property Zeilenbeschriftung As String
    Public Property ObjektTyp As String
    Public Property Inhalt As String


    Public Sub New()
        _Titel = ""
        _Identifier = ""
        _InputFile = ""
        _Typ=""
        _Datenbereich = ""
        _TabellenNummer = "0"
        _TabellenName = ""
        _SpaltenNummer = "0"
        _Spaltenüberschrift = ""
        _ZeilenNummer = "0"
        _Zeilenbeschriftung = ""
        _ObjektTyp = ""
        _Inhalt = ""
    End Sub

End Class
