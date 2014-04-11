Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices.Core
Imports System.Text.RegularExpressions
Module ExtensionMethods
    Sub New()
    End Sub

    ''' <summary>
    ''' Allows a neater ForEach iteration of the AutoCAD database.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="database"></param>
    ''' <param name="action"></param>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension> _
    Public Sub ForEach(Of T As Entity)(database As Database, action As Action(Of T))
        Using tr = database.TransactionManager.StartTransaction()
            ' Get the model space block table record
            Dim currentspace = DirectCast(tr.GetObject(BlockReplace.Active.Database.CurrentSpaceId, OpenMode.ForRead), BlockTableRecord)
            'Dim modelSpace = DirectCast(tr.GetObject(blockTable(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

            Dim theClass As Autodesk.AutoCAD.Runtime.RXClass = Autodesk.AutoCAD.Runtime.RXObject.GetClass(GetType(T))

            ' Loop through the entities in model space
            For Each objectId As ObjectId In currentspace
                ' Look for entities of the correct type
                If objectId.ObjectClass.IsDerivedFrom(theClass) Then
                    Dim entity = DirectCast(tr.GetObject(objectId, OpenMode.ForRead), T)

                    action(entity)
                End If
            Next
            tr.Commit()
        End Using
    End Sub
    ''' <summary>
    ''' From this page: http://stackoverflow.com/questions/1381097/regex-get-the-name-of-captured-groups-in-c-sharp
    ''' Answer number 2
    ''' </summary>
    ''' <param name="regex"></param>
    ''' <param name="input"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")> <System.Runtime.CompilerServices.Extension> _
    Public Function MatchNamedCaptures(regex As Regex, input As String) As Dictionary(Of String, String)
        Dim namedCaptureDictionary = New Dictionary(Of String, String)()
        Dim groups As GroupCollection = regex.Match(input).Groups
        Dim groupNames As String() = regex.GetGroupNames()
        For Each groupName As String In groupNames
            If groups(groupName).Captures.Count > 0 Then
                namedCaptureDictionary.Add(groupName, groups(groupName).Value)
            End If
        Next
        Return namedCaptureDictionary
    End Function

    ''' <summary>
    ''' Returns the maximal element of the given sequence, based on
    ''' the given projection.
    ''' copied from: https://code.google.com/p/morelinq/source/browse/MoreLinq/MaxBy.cs
    ''' </summary>
    ''' <remarks>
    ''' If more than one element has the maximal projected value, the first
    ''' one encountered will be returned. This overload uses the default comparer
    ''' for the projected type. This operator uses immediate execution, but
    ''' only buffers a single result (the current maximal element).
    ''' </remarks>
    ''' <typeparam name="TSource">Type of the source sequence</typeparam>
    ''' <typeparam name="TKey">Type of the projected element</typeparam>
    ''' <param name="source">Source sequence</param>
    ''' <param name="selector">Selector to use to pick the results to compare</param>
    ''' <returns>The maximal element, according to the projection.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/> or <paramref name="selector"/> is null</exception>
    ''' <exception cref="InvalidOperationException"><paramref name="source"/> is empty</exception>

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")> <System.Runtime.CompilerServices.Extension> _
    Public Function MaxBy(Of TSource, TKey)(source As IEnumerable(Of TSource), selector As Func(Of TSource, TKey)) As TSource
        Return source.MaxBy(selector, Comparer(Of TKey).[Default])
    End Function

    ''' <summary>
    ''' Returns the maximal element of the given sequence, based on
    ''' the given projection and the specified comparer for projected values. 
    ''' copied from: https://code.google.com/p/morelinq/source/browse/MoreLinq/MaxBy.cs
    ''' </summary>
    ''' <remarks>
    ''' If more than one element has the maximal projected value, the first
    ''' one encountered will be returned. This overload uses the default comparer
    ''' for the projected type. This operator uses immediate execution, but
    ''' only buffers a single result (the current maximal element).
    ''' </remarks>
    ''' <typeparam name="TSource">Type of the source sequence</typeparam>
    ''' <typeparam name="TKey">Type of the projected element</typeparam>
    ''' <param name="source">Source sequence</param>
    ''' <param name="selector">Selector to use to pick the results to compare</param>
    ''' <param name="comparer">Comparer to use to compare projected values</param>
    ''' <returns>The maximal element, according to the projection.</returns>
    ''' <exception cref="ArgumentNullException"><paramref name="source"/>, <paramref name="selector"/> 
    ''' or <paramref name="comparer"/> is null</exception>
    ''' <exception cref="InvalidOperationException"><paramref name="source"/> is empty</exception>

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")> <System.Runtime.CompilerServices.Extension> _
    Public Function MaxBy(Of TSource, TKey)(source As IEnumerable(Of TSource), selector As Func(Of TSource, TKey), comparer As IComparer(Of TKey)) As TSource
        If source Is Nothing Then
            Throw New ArgumentNullException("source")
        End If
        If selector Is Nothing Then
            Throw New ArgumentNullException("selector")
        End If
        If comparer Is Nothing Then
            Throw New ArgumentNullException("comparer")
        End If
        Using sourceIterator = source.GetEnumerator()
            If Not sourceIterator.MoveNext() Then
                Throw New InvalidOperationException("Sequence contains no elements")
            End If
            Dim max = sourceIterator.Current
            Dim maxKey = selector(max)
            While sourceIterator.MoveNext()
                Dim candidate = sourceIterator.Current
                Dim candidateProjected = selector(candidate)
                If comparer.Compare(candidateProjected, maxKey) > 0 Then
                    max = candidate
                    maxKey = candidateProjected
                End If
            End While
            Return max
        End Using
    End Function
#Region "Unused Code"


        ' ''' <summary>
        ' ''' Returns the transformation matrix from the ViewTableRecord DCS to WCS
        ' ''' </summary>
        ' ''' <param name="view"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        '<System.Runtime.CompilerServices.Extension> _
        'Public Function EyeToWorld(view As ViewTableRecord) As Matrix3d
        '    Return Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) * Matrix3d.Displacement(view.Target - Point3d.Origin) * Matrix3d.PlaneToWorld(view.ViewDirection)
        'End Function

        ' ''' <summary>
        ' ''' Returns the transformation matrix from WCS to the ViewTableRecord DCS
        ' ''' </summary>
        ' ''' <param name="view"></param>
        ' ''' <returns></returns>
        ' ''' <remarks></remarks>
        '<System.Runtime.CompilerServices.Extension> _
        'Public Function WorldToEye(view As ViewTableRecord) As Matrix3d
        '    Return view.EyeToWorld().Inverse()
        'End Function

        ' ''' <summary>
        ' ''' Process a zoom according to the extents3d in the current viewport
        ' ''' </summary>
        ' ''' <param name="ed"></param>
        ' ''' <param name="ext"></param>
        ' ''' <remarks></remarks>
        '<System.Runtime.CompilerServices.Extension> _
        'Public Sub Zoom(ed As Editor, ext As Extents3d)
        '    Using view As ViewTableRecord = ed.GetCurrentView()
        '        ext.TransformBy(view.WorldToEye())
        '        view.Width = ext.MaxPoint.X - ext.MinPoint.X
        '        view.Height = ext.MaxPoint.Y - ext.MinPoint.Y
        '        view.CenterPoint = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2.0, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2.0)
        '        ed.SetCurrentView(view)
        '    End Using
        'End Sub

        ' ''' <summary>
        ' ''' Process a zoom extents in the current viewport
        ' ''' </summary>
        ' ''' <param name="ed"></param>
        ' ''' <remarks></remarks>
        '<System.Runtime.CompilerServices.Extension> _
        'Public Sub ZoomExtents(ed As Editor)
        '    Dim db As Database = ed.Document.Database
        '    Dim ext As Extents3d = If(CShort(Application.GetSystemVariable("cvport")) = 1, New Extents3d(db.Pextmin, db.Pextmax), New Extents3d(db.Extmin, db.Extmax))
        '    ed.Zoom(ext)
        'End Sub
#End Region
End Module
