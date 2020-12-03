Public Class clsAppearanceDB

    Public Property id As Date

    Public Property listofAppearances As List(Of clsAppearance)

    Public Sub copyFrom(ByVal appearanceDef As SortedList(Of String, clsAppearance))

        For Each kvp As KeyValuePair(Of String, clsAppearance) In appearanceDef
            Me.listofAppearances.Add(kvp.Value)
        Next

    End Sub

    Public Sub copyto(ByRef appearanceDef As SortedList(Of String, clsAppearance))
        For Each appDef In Me.listofAppearances
            appearanceDef.Add(appDef.name, appDef)
        Next
    End Sub



    Public Sub New()
        _id = Date.MinValue
        _listofAppearances = New List(Of clsAppearance)
    End Sub



    'Public name As String
    'Public isMilestone As Boolean
    'Public FGcolor As Integer          'shp.Fill.ForeColor.RGB
    'Public BGcolor As Integer          'shp.Fill.BackColor.RGB
    'Public Rotation As Single          'shp.Rotation
    'Public Glowcolor As Integer               'shp.Glow.Color.RGB
    'Public Glowradius As Integer       'shp.Glow.Radius
    'Public ShadowFG As Integer                'shp.Shadow.ForeColor.RGB
    'Public ShadowTransp As Integer   'shp.Shadow.Transparency
    'Public shpType As Microsoft.Office.Core.MsoAutoShapeType 'shp.AutoShapeType
    'Public width As Single             'shp.Width
    'Public height As Single            'shp.Height
    'Public LineBGColor As Integer      'shp.Line.BackColor
    'Public LineFGColor As Integer      'shp.Line.ForeColor
    'Public LineWeight As Single        'shp.Line.Weight
    'Public hasText As Boolean          'shp.TextFrame2.hasText
    'Public TextMarginLeft As Single    'shp.TextFrame2.MarginLeft
    'Public TextMarginRight As Single   'shp.TextFrame2.MarginRight
    'Public TextMarginBottom As Single  'shp.TextFrame2.MarginBottom
    'Public TextMarginTop As Single     'shp.TextFrame2.MarginTop
    'Public TextWordWrap As Object      'shp.TextFrame2.WordWrap
    'Public TextVerticalAnchor As Object  'shp.TextFrame2.VerticalAnchor
    'Public TextHorizontalAnchor As Object 'shp.TextFrame2.HorizontalAnchor
    'Public TextRangeText As String        'shp.TextFrame2.TextRange.Text
    'Public TextRangeFontSize As Single  'shp.TextFrame2.TextRange.Font.Size
    'Public TextRangeFontFillFGColor As Integer 'shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB


    'Public Sub CopyTo(ByRef newApp As clsAppearance)

    '    With newApp

    '    End With

    'End Sub

    'Public Sub Copyfrom(ByVal app As clsAppearance)
    '    Me.name = ""
    '    Me.isMilestone = False
    '    '_form = Nothing
    '    Me.FGcolor = 0          'shp.Fill.ForeColor.RGB
    '    Me.BGcolor = 0          'shp.Fill.BackColor.RGB
    '    Me.Glowcolor = 0        'shp.Glow.Color.RGB
    '    Me.Glowradius = 0       'shp.Glow.Radius
    '    Me.ShadowFG = 0         'shp.Shadow.ForeColor.RGB
    '    Me.ShadowTransp = 0     'shp.Shadow.Transparency
    '    Me.shpType = Nothing    'shp.AutoShapeType
    '    Me.width = 0            'shp.Width
    '    Me.height = 0           'shp.Height
    '    Me.LineBGColor = 0      'shp.Line.BackColor
    '    Me.LineFGColor = 0      'shp.Line.ForeColor
    '    Me.LineWeight = 0       'shp.Line.Weight
    '    Me.TextMarginLeft = 0   'shp.TextFrame2.MarginLeft
    '    TextMarginRight = 0  'shp.TextFrame2.MarginRight
    '    TextMarginBottom = 0 'shp.TextFrame2.MarginBottom
    '    TextMarginTop = 0    'shp.TextFrame2.MarginTop
    '    TextWordWrap = Microsoft.Office.Core.MsoTriState.msoFalse     'shp.TextFrame2.WordWrap
    '    TextVerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle  'shp.TextFrame2.VerticalAnchor
    '    TextHorizontalAnchor = Microsoft.Office.Core.MsoHorizontalAnchor.msoAnchorCenter  'shp.TextFrame2.HorizontalAnchor
    '    TextRangeText = ""     'shp.TextFrame2.TextRange.Text
    '    TextRangeFontSize = CSng(awinSettings.fontsizeLegend)   'shp.TextFrame2.TextRange.Font.Size
    '    TextRangeFontFillFGColor = RGB(255, 255, 255) 'shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB

    'End Sub


    'Public Sub New()
    '    name = ""
    '    isMilestone = False
    '    '_form = Nothing
    '    FGcolor = 0          'shp.Fill.ForeColor.RGB
    '    BGcolor = 0          'shp.Fill.BackColor.RGB
    '    Glowcolor = 0        'shp.Glow.Color.RGB
    '    Glowradius = 0       'shp.Glow.Radius
    '    ShadowFG = 0         'shp.Shadow.ForeColor.RGB
    '    ShadowTransp = 0     'shp.Shadow.Transparency
    '    shpType = Nothing    'shp.AutoShapeType
    '    width = 0            'shp.Width
    '    height = 0           'shp.Height
    '    LineBGColor = 0      'shp.Line.BackColor
    '    LineFGColor = 0      'shp.Line.ForeColor
    '    LineWeight = 0       'shp.Line.Weight
    '    TextMarginLeft = 0   'shp.TextFrame2.MarginLeft
    '    TextMarginRight = 0  'shp.TextFrame2.MarginRight
    '    TextMarginBottom = 0 'shp.TextFrame2.MarginBottom
    '    TextMarginTop = 0    'shp.TextFrame2.MarginTop
    '    TextWordWrap = Microsoft.Office.Core.MsoTriState.msoFalse     'shp.TextFrame2.WordWrap
    '    TextVerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle  'shp.TextFrame2.VerticalAnchor
    '    TextHorizontalAnchor = Microsoft.Office.Core.MsoHorizontalAnchor.msoAnchorCenter  'shp.TextFrame2.HorizontalAnchor
    '    TextRangeText = ""     'shp.TextFrame2.TextRange.Text
    '    TextRangeFontSize = CSng(awinSettings.fontsizeLegend)   'shp.TextFrame2.TextRange.Font.Size
    '    TextRangeFontFillFGColor = RGB(255, 255, 255) 'shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    'End Sub

End Class
