Imports System.Diagnostics
Imports System.Windows.Forms
Imports ProjectBoardDefinitions
Public Class ucProperties

    ' nimmt den aktuell gültigen docLink auf 
    Private _documentsLink As String
    Private _myDocumentsLink As String

    Private _mediaLink As String
    Private _myMedialink As String

    Private _survLink As String
    Private _mySurvLink As String

    Private _3DLink As String
    Private _my3DLink As String

    ' nimmt auf , um welches Shape es sich aktuell handelt 
    Private _currentShape As PowerPoint.Shape

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        _documentsLink = ""
        _myDocumentsLink = ""

        _mediaLink = ""
        _myMedialink = ""

        _survLink = ""
        _mySurvLink = ""

        _3DLink = ""
        _my3DLink = ""

        _currentShape = Nothing


    End Sub

    ''' <summary>
    ''' setzt alle Picture Boxes und Container auf Visible bzw nicht visible, in Abhängigkeit von visibleStatus 
    ''' </summary>
    Friend Sub setLinksToVisible(ByVal visibleStatus As Boolean)

        stdLinks.Visible = visibleStatus
        myLinks.Visible = visibleStatus

        doclnk.Visible = visibleStatus
        mydoclnk.Visible = visibleStatus
        medialnk.Visible = visibleStatus
        mymedialnk.Visible = visibleStatus
        survlnk.Visible = visibleStatus
        mysurvlnk.Visible = visibleStatus
        dreiDlnk.Visible = visibleStatus
        mydreiDlnk.Visible = visibleStatus
    End Sub

    ''' <summary>
    ''' setzt die Link-Werte, die hinter den Pictureboxes für das betreffende Shape liegen .. 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    Friend Sub setLinkValues(ByVal tmpShape As PowerPoint.Shape)
        currentShape = tmpShape
        If Not IsNothing(tmpShape) Then
            ' jetzt zuweisen ... dabei werden, je nach Wert auch die Werte Enabled bzw. Wahl des Bildes gezeigt ..  
            Me.documentsLink = tmpShape.Tags.Item("DUC")
            Me.myDocumentsLink = tmpShape.Tags.Item("DUM")

            Me.mediaLink = tmpShape.Tags.Item("MUC")
            Me.myMediaLink = tmpShape.Tags.Item("MUM")

            Me.survLink = tmpShape.Tags.Item("SUC")
            Me.mySurvLink = tmpShape.Tags.Item("SUM")

            Me.threeDLink = tmpShape.Tags.Item("3UC")
            Me.my3DLink = tmpShape.Tags.Item("3UM")

        Else

            Me.documentsLink = ""
            Me.myDocumentsLink = ""

            Me.mediaLink = ""
            Me.myMediaLink = ""

            Me.survLink = ""
            Me.mySurvLink = ""

            Me.threeDLink = ""
            Me.my3DLink = ""

        End If

    End Sub

    ''' <summary>
    ''' liest schreibt den general_documents link; wird beim Report Erzeugen ausgehend vom Original Plan gesetzt
    ''' </summary>
    ''' <returns></returns>
    Friend Property documentsLink As String
        Get
            documentsLink = _documentsLink
        End Get
        Set(value As String)
            Try
                If Not IsNothing(value) Then
                    ' die Prüfung , ob es sich um eine valide URL handelt erfolgt vor diesem Aufruf ! 
                    If value <> "" Then
                        _documentsLink = value
                        doclnk.Image = My.Resources.documents
                        doclnk.Visible = True
                    Else
                        _documentsLink = ""
                        doclnk.Image = My.Resources.documents_plus
                        doclnk.Visible = False
                    End If
                Else
                    _documentsLink = ""
                    doclnk.Visible = False
                End If

            Catch ex As Exception
                _documentsLink = ""
                doclnk.Visible = False
            End Try

        End Set
    End Property

    ''' <summary>
    ''' liest, schreibt den myDocuments link 
    ''' kann vom Powerpoint Empfänger selber gesetzt werden 
    ''' </summary>
    ''' <returns></returns>
    Friend Property myDocumentsLink As String
        Get
            myDocumentsLink = _myDocumentsLink
        End Get
        Set(value As String)
            Try
                If Not IsNothing(value) Then
                    ' die Prüfung , ob es sich um eine valide URL handelt erfolgt vor diesem Aufruf ! 
                    If value <> "" Then
                        _myDocumentsLink = value
                        mydoclnk.Image = My.Resources.documents
                    Else
                        _myDocumentsLink = ""
                        mydoclnk.Image = My.Resources.documents_plus
                    End If
                Else
                    _myDocumentsLink = ""
                    mydoclnk.Image = My.Resources.documents_plus
                End If

            Catch ex As Exception
                _myDocumentsLink = ""
                mydoclnk.Image = My.Resources.documents_plus
            End Try
        End Set
    End Property

    ''' <summary>
    ''' liest/schreibt den Media Link, also Fotos, Videos, etc
    ''' </summary>
    ''' <returns></returns>
    Friend Property mediaLink As String
        Get
            mediaLink = _mediaLink
        End Get
        Set(value As String)

            Try
                If Not IsNothing(value) Then
                    ' die Prüfung , ob es sich um eine valide URL handelt erfolgt vor diesem Aufruf ! 
                    If value <> "" Then
                        _mediaLink = value
                        medialnk.Visible = True
                    Else
                        _mediaLink = ""
                        medialnk.Visible = False
                    End If
                Else
                    _mediaLink = ""
                    medialnk.Visible = False
                End If

            Catch ex As Exception
                _mediaLink = ""
                medialnk.Visible = False
            End Try

        End Set
    End Property


    ''' <summary>
    ''' liest/schreibt den MyMedia Link, also Fotos, Videos, etc die der PPT Nutzer verlinken will
    ''' </summary>
    ''' <returns></returns>
    Friend Property myMediaLink As String
        Get
            myMediaLink = _myMedialink
        End Get
        Set(value As String)

            Try
                If Not IsNothing(value) Then
                    ' die Prüfung , ob es sich um eine valide URL handelt erfolgt vor diesem Aufruf ! 
                    If value <> "" Then
                        _myMedialink = value
                        mymedialnk.Image = My.Resources.camera
                    Else
                        _myMedialink = ""
                        mymedialnk.Image = My.Resources.camera_plus
                    End If
                Else
                    _myMedialink = ""
                    mymedialnk.Image = My.Resources.camera_plus
                End If

            Catch ex As Exception
                _myMedialink = ""
                mymedialnk.Image = My.Resources.camera_plus
            End Try

        End Set
    End Property


    ''' <summary>
    ''' liest/schreibt den webcam / surveillance Link, webCams zu Baustellen etc.
    ''' </summary>
    ''' <returns></returns>
    Friend Property survLink As String
        Get
            survLink = _survLink
        End Get
        Set(value As String)
            Try
                If Not IsNothing(value) Then
                    ' die Prüfung , ob es sich um eine valide URL handelt erfolgt vor diesem Aufruf ! 
                    If value <> "" Then
                        _survLink = value
                        survlnk.Visible = True
                    Else
                        _survLink = ""
                        survlnk.Visible = False
                    End If
                Else
                    _survLink = ""
                    survlnk.Visible = False
                End If

            Catch ex As Exception
                _survLink = ""
                survlnk.Visible = False
            End Try

        End Set
    End Property

    ''' <summary>
    ''' liest/schreibt den webcam / surveillance Link, webCams zu Baustellen etc.
    ''' </summary>
    ''' <returns></returns>
    Friend Property mySurvLink As String
        Get
            mySurvLink = _mySurvLink
        End Get
        Set(value As String)

            Try
                If Not IsNothing(value) Then
                    ' die Prüfung , ob es sich um eine valide URL handelt erfolgt vor diesem Aufruf ! 
                    If value <> "" Then
                        _mySurvLink = value
                        mysurvlnk.Image = My.Resources.surveillance_camera
                    Else
                        _mySurvLink = ""
                        mysurvlnk.Image = My.Resources.surveillance_camera_plus
                    End If
                Else
                    _mySurvLink = ""
                    mysurvlnk.Image = My.Resources.surveillance_camera_plus
                End If

            Catch ex As Exception
                _mySurvLink = ""
                mysurvlnk.Image = My.Resources.surveillance_camera_plus
            End Try


        End Set
    End Property

    ''' <summary>
    ''' liest schreibt den 3DLink - zentral
    ''' </summary>
    ''' <returns></returns>
    Friend Property threeDLink As String
        Get
            threeDLink = _3DLink
        End Get
        Set(value As String)
            Try
                If Not IsNothing(value) Then
                    ' die Prüfung , ob es sich um eine valide URL handelt erfolgt vor diesem Aufruf ! 
                    If value <> "" Then
                        _3DLink = value
                        dreiDlnk.Visible = True
                    Else
                        _3DLink = ""
                        dreiDlnk.Visible = False
                    End If
                Else
                    _3DLink = ""
                    dreiDlnk.Visible = False
                End If

            Catch ex As Exception
                _3DLink = ""
                dreiDlnk.Visible = False
            End Try

        End Set
    End Property

    ''' <summary>
    ''' liest schreibt den persönlichen 3DLink
    ''' </summary>
    ''' <returns></returns>
    Friend Property my3DLink As String
        Get
            my3DLink = _my3DLink
        End Get
        Set(value As String)
            Try
                If Not IsNothing(value) Then
                    ' die Prüfung , ob es sich um eine valide URL handelt erfolgt vor diesem Aufruf ! 
                    If value <> "" Then
                        _my3DLink = value
                        mydreiDlnk.Image = My.Resources._3d
                    Else
                        _my3DLink = ""
                        mydreiDlnk.Image = My.Resources._3d_plus
                    End If
                Else
                    _my3DLink = ""
                    mydreiDlnk.Image = My.Resources._3d_plus
                End If

            Catch ex As Exception
                _my3DLink = value
                mydreiDlnk.Image = My.Resources._3d_plus
            End Try

        End Set
    End Property

    ''' <summary>
    ''' liest/schreibt das Shape, um das es gerade im Info Pane geht ...
    ''' </summary>
    ''' <returns></returns>
    Friend Property currentShape As PowerPoint.Shape
        Get
            currentShape = _currentShape
        End Get
        Set(value As PowerPoint.Shape)
            _currentShape = value
        End Set
    End Property

    Private Sub ucProperties_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        ' label resize
        eleName.MaximumSize = New Drawing.Size(Me.Width - eleName.Margin.Left - eleName.Margin.Right - eleName.Location.X, eleName.MaximumSize.Height)

    End Sub

    Private Sub ucProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If englishLanguage Then
            With Me
                .labelAmpel.Text = "Traffic Light:"
                .labelDate.Text = "Date:"
                .labelDeliver.Text = "Deliverables:"
                .labelRespons.Text = "Responsible:"
            End With
        Else
            With Me
                .labelAmpel.Text = "Ampel:"
                .labelDate.Text = "Datum:"
                .labelDeliver.Text = "Leistungsumfänge:"
                .labelRespons.Text = "Verantwortlich:"
            End With
        End If



    End Sub

    ''' <summary>
    ''' blendet die benötigten Darstellungs-Elemente ein bzw aus und positioniert dieses 
    ''' </summary>
    ''' <param name="isOn"></param>
    ''' <remarks></remarks>
    Public Sub symbolMode(ByVal isOn As Boolean)
        Dim tmpLocation As New System.Drawing.Point
        Dim tmpSize As New System.Drawing.Size

        If isOn Then
            ' der Symbol Mode
            With labelDate
                .Visible = False
            End With

            With labelRespons
                .Visible = False
            End With

            With labelAmpel
                .Visible = False
            End With

            With eleAmpel
                .Visible = False
            End With

            With eleAmpelText
                .Visible = True
                tmpLocation.X = 5
                tmpLocation.Y = 52
                .Location = tmpLocation
                .BorderStyle = Windows.Forms.BorderStyle.None
                tmpSize.Height = 400
                'tmpSize.Width = 276
                .Size = tmpSize
            End With

            With labelDeliver
                .Visible = False
            End With

            With eleDeliverables
                .Visible = False
            End With

        Else
            ' der Normal-Mode
            With labelDate
                .Visible = True
                tmpLocation.X = 5
                tmpLocation.Y = 52
                .Location = tmpLocation
            End With

            With labelRespons
                .Visible = True
            End With

            With labelAmpel
                .Visible = True
            End With

            With eleAmpel
                .Visible = True
            End With

            With eleAmpelText
                .Visible = True
                tmpLocation.X = 10
                tmpLocation.Y = 144
                .Location = tmpLocation
                .BorderStyle = Windows.Forms.BorderStyle.FixedSingle
                tmpSize.Height = 139
                'tmpSize.Width = 276
                .Size = tmpSize
            End With

            With labelDeliver
                .Visible = True
            End With

            With eleDeliverables
                .Visible = True
            End With
        End If
    End Sub

    Private Sub eleAmpelText_MouseDoubleClick(sender As Object, e As Windows.Forms.MouseEventArgs) Handles eleAmpelText.MouseDoubleClick

    End Sub


    Private Sub eleAmpelText_TextChanged(sender As Object, e As EventArgs) Handles eleAmpelText.TextChanged

    End Sub

    ''' <summary>
    ''' geht auf die entsprechende URL - muss später ersetzt werden durch Rest Server Aufruf ...
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub doclnk_Click(sender As Object, e As EventArgs) Handles doclnk.Click
        If documentsLink.Length > 0 Then
            If isValidURL(documentsLink) Then
                Process.Start(documentsLink)
            Else
                Call MsgBox("invalid Url: " & documentsLink)
            End If
        End If

    End Sub

    Private Sub medialnk_Click(sender As Object, e As EventArgs) Handles medialnk.Click
        If mediaLink.Length > 0 Then
            If isValidURL(mediaLink) Then
                Process.Start(mediaLink)
            Else
                Call MsgBox("invalid Url: " & mediaLink)
            End If
        End If
    End Sub

    Private Sub survlnk_Click(sender As Object, e As EventArgs) Handles survlnk.Click
        If survLink.Length > 0 Then
            If isValidURL(survLink) Then
                Process.Start(survLink)
            Else
                Call MsgBox("invalid Url: " & survLink)
            End If
        End If
    End Sub

    Private Sub dreiDlnk_Click(sender As Object, e As EventArgs) Handles dreiDlnk.Click
        If threeDLink.Length > 0 Then
            If isValidURL(threeDLink) Then
                Process.Start(threeDLink)
            Else
                Call MsgBox("invalid Url: " & threeDLink)
            End If
        End If
    End Sub

   
    Private Sub mydoclnk_MouseClick(sender As Object, e As MouseEventArgs) Handles mydoclnk.MouseClick
        If Me.myDocumentsLink.Length > 0 And e.Button = MouseButtons.Left Then
            If isValidURL(Me.myDocumentsLink) Then
                Process.Start(Me.myDocumentsLink)
            Else
                Call MsgBox("invalid Url: " & Me.myDocumentsLink)
            End If
        Else
            ' das Formular zur Eingabe aufrufen 
            Dim linkForm As New frmEditLink
            linkForm.linkValue.Text = myDocumentsLink
            linkForm.titleExtension = " for documents"

            Dim res As Windows.Forms.DialogResult = linkForm.ShowDialog
            If res = Windows.Forms.DialogResult.OK Then

                currentShape.Tags.Add("DUM", linkForm.linkValue.Text)
                Me.myDocumentsLink = linkForm.linkValue.Text

            ElseIf res = Windows.Forms.DialogResult.No Then

                currentShape.Tags.Delete("DUM")
                Me.myDocumentsLink = ""

            End If

        End If
    End Sub

    Private Sub mymedialnk_MouseClick(sender As Object, e As MouseEventArgs) Handles mymedialnk.MouseClick
        If Me.myMediaLink.Length > 0 And e.Button = MouseButtons.Left Then
            If isValidURL(Me.myMediaLink) Then
                Process.Start(Me.myMediaLink)
            Else
                Call MsgBox("invalid Url: " & Me.myMediaLink)
            End If
        Else
            ' das Formular zur Eingabe aufrufen 
            Dim linkForm As New frmEditLink
            linkForm.linkValue.Text = myMediaLink
            linkForm.titleExtension = " for fotos & videos"

            Dim res As Windows.Forms.DialogResult = linkForm.ShowDialog
            If res = Windows.Forms.DialogResult.OK Then

                currentShape.Tags.Add("MUM", linkForm.linkValue.Text)
                Me.myMediaLink = linkForm.linkValue.Text

            ElseIf res = Windows.Forms.DialogResult.No Then

                currentShape.Tags.Delete("MUM")
                Me.myMediaLink = ""

            End If

        End If
    End Sub

    Private Sub mysurvlnk_MouseClick(sender As Object, e As MouseEventArgs) Handles mysurvlnk.MouseClick
        If Me.mySurvLink.Length > 0 And e.Button = MouseButtons.Left Then
            If isValidURL(Me.mySurvLink) Then
                Process.Start(Me.mySurvLink)
            Else
                Call MsgBox("invalid Url: " & Me.mySurvLink)
            End If
        Else
            ' das Formular zur Eingabe aufrufen 
            Dim linkForm As New frmEditLink
            linkForm.linkValue.Text = mySurvLink
            linkForm.titleExtension = " for webCams"

            Dim res As Windows.Forms.DialogResult = linkForm.ShowDialog
            If res = Windows.Forms.DialogResult.OK Then

                currentShape.Tags.Add("SUM", linkForm.linkValue.Text)
                Me.mySurvLink = linkForm.linkValue.Text

            ElseIf res = Windows.Forms.DialogResult.No Then

                currentShape.Tags.Delete("SUM")
                Me.mySurvLink = ""

            End If

        End If
    End Sub

    Private Sub mydreiDlnk_MouseClick(sender As Object, e As MouseEventArgs) Handles mydreiDlnk.MouseClick
        If Me.my3DLink.Length > 0 And e.Button = MouseButtons.Left Then
            If isValidURL(Me.my3DLink) Then
                Process.Start(Me.my3DLink)
            Else
                Call MsgBox("invalid Url: " & Me.my3DLink)
            End If
        Else
            ' das Formular zur Eingabe aufrufen 
            Dim linkForm As New frmEditLink
            linkForm.linkValue.Text = my3DLink
            linkForm.titleExtension = " for 3D-models"


            Dim res As Windows.Forms.DialogResult = linkForm.ShowDialog
            If res = Windows.Forms.DialogResult.OK Then

                currentShape.Tags.Add("3UM", linkForm.linkValue.Text)
                Me.my3DLink = linkForm.linkValue.Text

            ElseIf res = Windows.Forms.DialogResult.No Then

                currentShape.Tags.Delete("3UM")
                Me.my3DLink = ""

            End If

        End If
    End Sub

    ''' <summary>
    ''' zeigt den Link als Tooltipp 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub doclnk_MouseHover(sender As Object, e As EventArgs) Handles doclnk.MouseHover
        Dim ttmsg As String = ""
        If documentsLink.Length > 0 Then
            ttmsg = documentsLink
            ToolTip1.Show(ttmsg, doclnk, 2000)
        End If

    End Sub

    Private Sub medialnk_MouseHover(sender As Object, e As EventArgs) Handles medialnk.MouseHover
        Dim ttmsg As String = ""
        If mediaLink.Length > 0 Then
            ttmsg = mediaLink
            ToolTip1.Show(ttmsg, medialnk, 2000)
        End If
    End Sub

    Private Sub survlnk_MouseHover(sender As Object, e As EventArgs) Handles survlnk.MouseHover
        Dim ttmsg As String = ""
        If survLink.Length > 0 Then
            ttmsg = survLink
            ToolTip1.Show(ttmsg, survlnk, 2000)
        End If
    End Sub

    Private Sub dreiDlnk_MouseHover(sender As Object, e As EventArgs) Handles dreiDlnk.MouseHover
        Dim ttmsg As String = ""
        If threeDLink.Length > 0 Then
            ttmsg = threeDLink
            ToolTip1.Show(ttmsg, dreiDlnk, 2000)
        End If
    End Sub

    Private Sub mydoclnk_MouseHover(sender As Object, e As EventArgs) Handles mydoclnk.MouseHover

        Dim ttmsg As String = "click to define link"
        If myDocumentsLink.Length > 0 Then
            ttmsg = Me.myDocumentsLink & vbLf & "right click to change"
        End If
        ToolTip1.Show(ttmsg, mydoclnk, 2000)
    End Sub

    Private Sub mymedialnk_MouseHover(sender As Object, e As EventArgs) Handles mymedialnk.MouseHover
        Dim ttmsg As String = "click to define link"
        If myMediaLink.Length > 0 Then
            ttmsg = Me.myMediaLink & vbLf & "right click to change"
        End If
        ToolTip1.Show(ttmsg, mymedialnk, 2000)
    End Sub

    Private Sub mysurvlnk_MouseHover(sender As Object, e As EventArgs) Handles mysurvlnk.MouseHover
        Dim ttmsg As String = "click to define link"
        If mySurvLink.Length > 0 Then
            ttmsg = Me.mySurvLink & vbLf & "right click to change"
        End If
        ToolTip1.Show(ttmsg, mysurvlnk, 2000)
    End Sub

    Private Sub mydreiDlnk_MouseHover(sender As Object, e As EventArgs) Handles mydreiDlnk.MouseHover
        Dim ttmsg As String = "click to define link"
        If my3DLink.Length > 0 Then
            ttmsg = Me.my3DLink & vbLf & "right click to change"
        End If
        ToolTip1.Show(ttmsg, mydreiDlnk, 2000)
    End Sub

End Class
