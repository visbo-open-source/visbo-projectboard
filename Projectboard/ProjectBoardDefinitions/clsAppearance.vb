Imports xlNS = Microsoft.Office.Interop.Excel

''' <summary>
''' beschreibt, wie ein Element dieser Darstellungsklasse beschrieben werden soll  
''' </summary>
''' <remarks></remarks>
Public Class clsAppearance

    Public Property name As String
    Public Property isMilestone As Boolean
    'Public Property form As xlNS.Shape
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
    Public Property hasText As Boolean          'shp.TextFrame2.hasText
    Public Property TextMarginLeft As Single    'shp.TextFrame2.MarginLeft
    Public Property TextMarginRight As Single   'shp.TextFrame2.MarginRight
    Public Property TextMarginBottom As Single  'shp.TextFrame2.MarginBottom
    Public Property TextMarginTop As Single     'shp.TextFrame2.MarginTop
    Public Property TextWordWrap As Object      'shp.TextFrame2.WordWrap
    Public Property TextVerticalAnchor As Object  'shp.TextFrame2.VerticalAnchor
    Public Property TextHorizontalAnchor As Object 'shp.TextFrame2.HorizontalAnchor
    Public Property TextRangeText As String        'shp.TextFrame2.TextRange.Text
    Public Property TextRangeFontSize As Single  'shp.TextFrame2.TextRange.Font.Size
    Public Property TextRangeFontFillFGColor As Integer 'shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB




    Public Sub New()
        _name = ""
        _isMilestone = False
        '_form = Nothing
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
        _TextMarginLeft = 0   'shp.TextFrame2.MarginLeft
        _TextMarginRight = 0  'shp.TextFrame2.MarginRight
        _TextMarginBottom = 0 'shp.TextFrame2.MarginBottom
        _TextMarginTop = 0    'shp.TextFrame2.MarginTop
        _TextWordWrap = Microsoft.Office.Core.MsoTriState.msoFalse     'shp.TextFrame2.WordWrap
        _TextVerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle  'shp.TextFrame2.VerticalAnchor
        _TextHorizontalAnchor = Microsoft.Office.Core.MsoHorizontalAnchor.msoAnchorCenter  'shp.TextFrame2.HorizontalAnchor
        _TextRangeText = ""     'shp.TextFrame2.TextRange.Text
        _TextRangeFontSize = CSng(awinSettings.fontsizeLegend)   'shp.TextFrame2.TextRange.Font.Size
        _TextRangeFontFillFGColor = RGB(255, 255, 255) 'shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    End Sub
End Class
