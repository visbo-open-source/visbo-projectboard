Public Class clsHierarchy
    Private _allNodes As SortedList(Of String, clsHierarchyNode)

    ''' <summary>
    ''' fügt der Hierarchy einen Knoten hinzu
    ''' </summary>
    ''' <param name="elemNode"></param>
    ''' <remarks></remarks>
    Public Sub addNode(ByRef elemNode As clsHierarchyNode, ByVal elemKey As String)

        Dim parentNode As clsHierarchyNode
        Dim parentNodeKey As String = elemNode.parentNodeKey


        ' jetzt wird der Parent-Node bestimmt , sofern er existiert 
        If parentNodeKey.Length > 0 Then
            Try
                parentNode = _allNodes.Item(parentNodeKey)
            Catch ex As Exception
                Throw New Exception(parentNodeKey & " existiert nicht ")
            End Try
        Else
            parentNode = Nothing
        End If

        ' jetzt den Parent Key eintragen 
        elemNode.parentNodeKey = parentNodeKey
        _allNodes.Add(elemKey, elemNode)

        ' jetzt das Child im Parent Node verankern 
        If Not IsNothing(parentNode) Then
            parentNode.addChild(elemKey)
        End If



    End Sub

    ''' <summary>
    ''' berechnet den Unique Namen ElemKey für den gegebenen elemName und Angabe, ob Meilenstein oder nicht 
    ''' findet auf jeden Fall einen Namen, der in der sortedList noch nicht enthalten ist  
    ''' </summary>
    ''' <param name="elemName"></param>
    ''' <param name="isMilestone"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property findUniqueElemKey(ByVal elemName As String, ByVal isMilestone As Boolean) As String
        Get
            Dim elemKey As String
            Dim lfdNr As Integer = 2

            
            elemKey = calcHryElemKey(elemName, isMilestone)

            If _allNodes.ContainsKey(elemKey) Then

                elemKey = calcHryElemKey(elemName, isMilestone, lfdNr)

                Do While _allNodes.ContainsKey(elemKey)
                    lfdNr = lfdNr + 1
                    elemKey = calcHryElemKey(elemName, isMilestone, lfdNr)
                Loop

            End If

            findUniqueElemKey = elemKey

        End Get
    End Property


    

    ''' <summary>
    ''' gibt an, ob dieser Schlüssel bereits in der Liste vorhanden ist
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsKey(ByVal uniqueID As String) As Boolean
        Get
            containsKey = _allNodes.ContainsKey(uniqueID)
        End Get
    End Property

    ''' <summary>
    ''' gibt den Phasen Index in der Liste der Phasen zurück 
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getIndexOfElem(ByVal uniqueID As String) As Integer
        Get
            If _allNodes.ContainsKey(uniqueID) Then
                getIndexOfElem = _allNodes.Item(uniqueID).indexOfElem
            Else
                getIndexOfElem = -1
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt den Hierarchie Knoten zurück , der die angegebene uniqueID hat
    ''' VORSICHT: liefert Nothing zurück , wenn die uniqueID nicht existiert 
    ''' </summary>
    ''' <param name="uniqueID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property item(ByVal uniqueID As String) As clsHierarchyNode
        Get
            If _allNodes.ContainsKey(uniqueID) Then
                item = _allNodes.Item(uniqueID)
            Else
                item = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste von Phasen-Indices zurück; jeder Index bezeichnet eine Phase, die den Elem-Namen trägt 
    ''' und die mit der angegeben Hierarchie übereinstimmt; Hierarchie kann  
    ''' wenn ein 1-elementiger Array zurückgegeben wird, dessen Wert = 0 ist, so existiert diese Phase nicht  
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseIndices(ByVal name As String, ByVal breadcrumb As String) As Integer()
        Get
            Dim phaseIndices() As Integer
            Dim first As Integer, last As Integer
            Dim elemID As String = calcHryElemKey(name, False)
            Dim i As Integer, k As Integer

            ' da die Liste sortiert ist, reicht es den Index des ersten und den Index des letzten Elements zu bestimmen 

            ReDim phaseIndices(0)
            If _allNodes.ContainsKey(elemID) Then
                first = _allNodes.IndexOfKey(calcHryElemKey(name, False))
                i = 1
                Do While elemNameOfElemID(_allNodes.ElementAt(first + i).Key) = name
                    i = i + 1
                Loop
                last = first + i - 1

                If breadcrumb = "" Then
                    ReDim phaseIndices(last - first)
                    For i = 0 To last - first
                        phaseIndices(i) = _allNodes.ElementAt(first + i).Value.indexOfElem
                    Next
                Else
                    Dim tmpIndices(last - first) As Integer
                    Dim identicalBC As Boolean = True
                    Dim bcLevels() As String
                    bcLevels = breadcrumb.Split((New Char() {CChar("#")}), 20)
                    Dim anzLevel As Integer = bcLevels.Length
                    k = 0
                    For i = 0 To last - first
                        Dim vglbreadCrumb As String = Me.getBreadCrumb(_allNodes.ElementAt(first + i).Key, anzLevel)
                        If vglbreadCrumb = breadcrumb Then
                            tmpIndices(k) = _allNodes.ElementAt(first + i).Value.indexOfElem
                            k = k + 1
                        End If
                        If k > 0 Then
                            ' es wurde mindestens ein Element gefunden 
                            k = k - 1
                            ReDim phaseIndices(k)
                            For ii As Integer = 0 To k
                                phaseIndices(ii) = tmpIndices(ii)
                            Next
                        End If
                    Next


                End If

            End If

            getPhaseIndices = phaseIndices

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste von Meilenstein-Indices zurück;  Index(0,) bezeichnet die Phasen-Nummer, Index(1,) bezeichnet die Meilenstein-Nummer 
    ''' es werden nur die Elemente zurückgegeben,  die mit der angegeben Hierarchie übereinstimmen; Breadcrumb kann leer sein, dann wird alles gesucht   
    ''' wenn ein 2-elementiger Array zurückgegeben wird, dessen Wert jeweils 0 ist, so existiert dieser Meilenstein nicht   
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="breadcrumb"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneIndices(ByVal name As String, ByVal breadcrumb As String) As Integer(,)
        Get
            Dim milestoneIndices(,) As Integer
            Dim first As Integer, last As Integer
            Dim elemID As String = calcHryElemKey(name, True)
            Dim i As Integer, k As Integer
            Dim phaseID As String

            ' da die Liste sortiert ist, reicht es den Index des ersten und den Index des letzten Elements zu bestimmen 

            ReDim milestoneIndices(1, 0)
            If _allNodes.ContainsKey(elemID) Then
                first = _allNodes.IndexOfKey(elemID)
                i = 1
                Do While elemNameOfElemID(_allNodes.ElementAt(first + i).Key) = name
                    i = i + 1
                Loop
                last = first + i - 1

                If breadcrumb = "" Then
                    ReDim milestoneIndices(1, last - first)
                    For i = 0 To last - first
                        phaseID = _allNodes.ElementAt(first + i).Value.parentNodeKey
                        milestoneIndices(0, i) = Me.getIndexOfElem(phaseID)
                        milestoneIndices(1, i) = _allNodes.ElementAt(first + i).Value.indexOfElem
                    Next
                Else
                    Dim tmpIndices(1, last - first) As Integer
                    Dim identicalBC As Boolean = True
                    Dim bcLevels() As String
                    bcLevels = breadcrumb.Split((New Char() {CChar("#")}), 20)
                    Dim anzLevel As Integer = bcLevels.Length
                    k = 0
                    For i = 0 To last - first
                        Dim vglbreadCrumb As String = Me.getBreadCrumb(_allNodes.ElementAt(first + i).Key, anzLevel)
                        If vglbreadCrumb = breadcrumb Then
                            phaseID = _allNodes.ElementAt(first + i).Value.parentNodeKey
                            tmpIndices(0, k) = Me.getIndexOfElem(phaseID)
                            tmpIndices(1, k) = _allNodes.ElementAt(first + i).Value.indexOfElem
                            k = k + 1
                        End If
                        If k > 0 Then
                            ' es wurde mindestens ein Element gefunden 
                            k = k - 1
                            ReDim milestoneIndices(1, k)
                            For ii As Integer = 0 To k
                                milestoneIndices(0, ii) = tmpIndices(0, ii)
                                milestoneIndices(1, ii) = tmpIndices(1, ii)
                            Next
                        End If
                    Next


                End If

            End If

            getMilestoneIndices = milestoneIndices

        End Get
    End Property


    ''' <summary>
    ''' gibt für die angegebene elemID den Breadcrumb zurück
    ''' die einzelnen Ebenen werden mit # voneinander getrennt
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <param name="ebene">100: die gesamte Hierarchie
    ''' 1, 2, ..: soviele Stufen wie angegeben sind; 
    ''' wenn mehr Stufen angegeben sind als vorhanden, wird bei . (root) abgebrochen</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBreadCrumb(ByVal elemID As String, Optional ByVal ebene As Integer = 100) As String
        Get
            Dim tmpBreadCrumb = ""
            Dim tmpEbene As Integer = 1
            Dim rootReached As Boolean = False
            Dim currentElemID As String = elemID
            Dim parentID As String = ""
            Dim rootkey As String = calcHryElemKey(".", False)

            ' sicherstellen, dass ebene keinen blödsinningen Wert hat ;
            ' wenn der Wert < 1 ist , wird automatisch auf 1 gesetzt 
            If ebene < 1 Then
                ebene = 1
            End If

            Dim ok As Boolean = True

            If elemID <> rootkey Then

                Do While tmpEbene <= ebene And Not rootReached And ok

                    If Me._allNodes.ContainsKey(currentElemID) Then

                        parentID = Me._allNodes.Item(currentElemID).parentNodeKey
                        If parentID = "" Then
                            rootReached = True
                        End If
                        If Me._allNodes.ContainsKey(parentID) Then

                            If tmpEbene = 1 Then
                                tmpBreadCrumb = Me._allNodes.Item(parentID).elemName
                            Else
                                tmpBreadCrumb = Me._allNodes.Item(parentID).elemName & "#" & tmpBreadCrumb
                            End If
                            currentElemID = parentID

                        Else
                            ok = False
                            If tmpEbene = 1 Then
                                tmpBreadCrumb = "?"
                            ElseIf Not rootReached Then
                                tmpBreadCrumb = "?" & "#" & tmpBreadCrumb
                            End If
                        End If

                    Else
                        ok = False
                        If tmpEbene = 1 Then
                            tmpBreadCrumb = "?"
                        Else
                            tmpBreadCrumb = "?" & "#" & tmpBreadCrumb
                        End If
                    End If


                    tmpEbene = tmpEbene + 1

                    
                Loop

            Else
                tmpBreadCrumb = ""
            End If

            getBreadCrumb = tmpBreadCrumb

        End Get
    End Property
    Public Sub New()
        _allNodes = New SortedList(Of String, clsHierarchyNode)
    End Sub


End Class
