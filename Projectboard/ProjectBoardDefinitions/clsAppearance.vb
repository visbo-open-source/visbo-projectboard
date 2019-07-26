Imports xlNS = Microsoft.Office.Interop.Excel

''' <summary>
''' beschreibt, wie ein Element dieser Darstellungsklasse beschrieben werden soll  
''' </summary>
''' <remarks></remarks>
Public Class clsAppearance

    Public Property name As String
    Public Property isMilestone As Boolean
    Public Property form As xlNS.Shape
    Public Property FGcolor As Integer          'shp.Fill.ForeColor.RGB
    Public Property BGcolor As Integer          'shp.Fill.BackColor.RGB
    Public Property Rotation As Single
    Public Property Glowcolor As Integer        'shp.Glow.Color.RGB
    Public Property Glowradius As Integer       'shp.Glow.Radius
    Public Property ShadowFG As Integer         'shp.Shadow.ForeColor.RGB
    Public Property ShadowTransp As Integer     'shp.Shadow.Transparency
    Public Property shpType As Microsoft.Office.Core.MsoAutoShapeType 'shp.AutoShapeType
    Public Property width As Single             'shp.Width
    Public Property height As Single            'shp.Height
    Public Property LineBGColor As Integer      'shp.Line.BackColor
    Public Property LineFGColor As Integer      'shp.Line.ForeColor
    Public Property LineWeight As Single        'shp.Line.Weight




    Public Sub New()
        _name = ""
        _isMilestone = False
        _form = Nothing
        _FGcolor = 0          'shp.Fill.ForeColor.RGB
        _BGcolor = 0          'shp.Fill.BackColor.RGB
        _Glowcolor = 0        'shp.Glow.Color.RGB
        _Glowradius = 0       'shp.Glow.Radius
        _ShadowFG = 0         'shp.Shadow.ForeColor.RGB
        _ShadowTransp = 0     'shp.Shadow.Transparency
        _shpType = Nothing    'shp.AutoShapeType
        _width = 0            'shp.Width
        _height = 0           'shp.Height
        _LineBGColor = 0      'shp.Line.BackColor
        _LineFGColor = 0      'shp.Line.ForeColor
        _LineWeight = 0       'shp.Line.Weight
    End Sub
End Class
