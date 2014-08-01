Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices.Core
Imports System.Text.RegularExpressions
Imports System.Collections
Imports System.Collections.Generic
Imports System.Linq

'Imports acadApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
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

    ''' <summary>
    ''' Synchronize BlockReference with the BlockDefinition
    ''' Copied from here: https://sites.google.com/site/bushmansnetlaboratory/moi-zametki/attsynch (via http://theSwamp.org)
    ''' </summary>
    ''' <param name="btr">table Entry for determining the block</param>
    ''' <param name="directOnly">Look at the top level or nested</param>
    ''' <param name="removeSuperfluous">Remove unecessary attributes</param>
    ''' <param name="setAttDefValues">Assign the current value or default.</param>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension> _
    Public Sub AttSync(btr As BlockTableRecord, directOnly As Boolean, removeSuperfluous As Boolean, setAttDefValues As Boolean)
        Dim db As Database = btr.Database
        Using wdb As New WorkingDatabaseSwitcher(db)
            Using t As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = DirectCast(t.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

                'get all the attributedefinitions of a block definition
                Dim attdefs As IEnumerable(Of AttributeDefinition) =
                    btr.Cast(Of ObjectId)().Where(Function(n) n.ObjectClass.Name = _
                                                      "AcDbAttributeDefinition").[Select](Function(n) _
                                                                                              DirectCast(t.GetObject(n, OpenMode.ForRead), AttributeDefinition)).Where(Function(n) Not n.Constant) 'exclude constant attributes, because AttributeReferences are not created for them.
                'loop through all occurences of a blockdefinition
                For Each brId As ObjectId In btr.GetBlockReferenceIds(directOnly, False)
                    Dim br As BlockReference = DirectCast(t.GetObject(brId, OpenMode.ForWrite), BlockReference)

                    'check for matching names
                    If br.Name <> btr.Name Then
                        Continue For
                    End If

                    'get attributereferences of the blockreference
                    Dim attrefs As IEnumerable(Of AttributeReference) = br.AttributeCollection.Cast(Of ObjectId)().[Select](Function(n) DirectCast(t.GetObject(n, OpenMode.ForWrite), AttributeReference))

                    'tags existing attributes in the attributedefinition entry
                    Dim definitiontags As IEnumerable(Of String) = attdefs.[Select](Function(n) n.Tag)
                    'collects tags existing attributes in the attributereference entry
                    Dim referencetags As IEnumerable(Of String) = attrefs.[Select](Function(n) n.Tag)
                    'remove superfluous attributes
                    If removeSuperfluous Then
                        For Each attref As AttributeReference In attrefs.Where(Function(n) referencetags.Except(definitiontags).Contains(n.Tag))
                            attref.[Erase](True)
                        Next
                    End If

                    'synchronize existing attributes and properties with their definitions.
                    For Each attref As AttributeReference In attrefs.Where(Function(n) definitiontags.Join(referencetags, Function(a) a, Function(b) b, Function(a, b) a).Contains(n.Tag))
                        Dim ad As AttributeDefinition = attdefs.First(Function(n) n.Tag = attref.Tag)

                        'Method SetAttributeFromBlock, we use later in the code, resets
                        'current value of the multiline attribute. So remember this value,
                        'to restore it immediately after calling SetAttributeFromBlock.
                        If attref.IsMTextAttribute Then
                            Dim value As String = attref.MTextAttribute.Text
                            Dim newValue() As String = value.Split(New Char() {Chr(13)}, StringSplitOptions.RemoveEmptyEntries)
                            attref.SetAttributeFromBlock(ad, br.BlockTransform)
                            attref.TextString = String.Join(vbCrLf, newValue)
                        Else
                            Dim val As String = attref.TextString
                            attref.SetAttributeFromBlock(ad, br.BlockTransform)
                            'restore attribute value
                            attref.TextString = val
                        End If

                        'If you want to - set to the default value of the attribute
                        If setAttDefValues Then
                            attref.TextString = ad.TextString
                        End If

                        attref.AdjustAlignment(db)
                    Next

                    'If the occurrence unit no desired attributes - create them
                    Dim attdefsNew As IEnumerable(Of AttributeDefinition) = attdefs.Where(Function(n) definitiontags.Except(referencetags).Contains(n.Tag))

                    For Each ad As AttributeDefinition In attdefsNew
                        Dim attref As New AttributeReference()
                        attref.SetAttributeFromBlock(ad, br.BlockTransform)
                        attref.AdjustAlignment(db)
                        br.AttributeCollection.AppendAttribute(attref)
                        t.AddNewlyCreatedDBObject(attref, True)
                    Next
                Next
                btr.UpdateAnonymousBlocks()
                t.Commit()
            End Using
            'If this is a dynamic block
            If btr.IsDynamicBlock Then
                Using t As Transaction = db.TransactionManager.StartTransaction()
                    For Each id As ObjectId In btr.GetAnonymousBlockIds()
                        Dim _btr As BlockTableRecord = DirectCast(t.GetObject(id, OpenMode.ForWrite), BlockTableRecord)

                        'Get all the attribute definitions from the original block definition
                        Dim attdefs As IEnumerable(Of AttributeDefinition) = btr.Cast(Of ObjectId)().Where(Function(n) n.ObjectClass.Name = "AcDbAttributeDefinition").[Select](Function(n) DirectCast(t.GetObject(n, OpenMode.ForRead), AttributeDefinition))

                        'Get all the attribute definitions from the definition of anonymous block
                        Dim attdefs2 As IEnumerable(Of AttributeDefinition) = _btr.Cast(Of ObjectId)().Where(Function(n) n.ObjectClass.Name = "AcDbAttributeDefinition").[Select](Function(n) DirectCast(t.GetObject(n, OpenMode.ForWrite), AttributeDefinition))

                        'attribute definitions of anonymous blocks to synchronize
                        'attribute definitions of the main unit

                        'Tags existing attribute definitions
                        Dim dtags As IEnumerable(Of String) = attdefs.[Select](Function(n) n.Tag)
                        Dim dtags2 As IEnumerable(Of String) = attdefs2.[Select](Function(n) n.Tag)

                        '1. Remove unnecessary
                        For Each attdef As AttributeDefinition In attdefs2.Where(Function(n) Not dtags.Contains(n.Tag))
                            attdef.[Erase](True)
                        Next

                        '2. Sync existing
                        For Each attdef As AttributeDefinition In attdefs.Where(Function(n) dtags.Join(dtags2, Function(a) a, Function(b) b, Function(a, b) a).Contains(n.Tag))
                            Dim ad As AttributeDefinition = attdefs2.First(Function(n) n.Tag = attdef.Tag)
                            ad.Position = attdef.Position
                            ad.TextStyleId = attdef.TextStyleId
                            'If you want to - set to the default value of the attribute
                            If setAttDefValues Then
                                ad.TextString = attdef.TextString
                            End If

                            ad.Tag = attdef.Tag
                            ad.Prompt = attdef.Prompt

                            ad.LayerId = attdef.LayerId
                            ad.Rotation = attdef.Rotation
                            ad.LinetypeId = attdef.LinetypeId
                            ad.LineWeight = attdef.LineWeight
                            ad.LinetypeScale = attdef.LinetypeScale
                            ad.Annotative = attdef.Annotative
                            ad.Color = attdef.Color
                            ad.Height = attdef.Height
                            ad.HorizontalMode = attdef.HorizontalMode
                            ad.Invisible = attdef.Invisible
                            ad.IsMirroredInX = attdef.IsMirroredInX
                            ad.IsMirroredInY = attdef.IsMirroredInY
                            ad.Justify = attdef.Justify
                            ad.LockPositionInBlock = attdef.LockPositionInBlock
                            ad.MaterialId = attdef.MaterialId
                            ad.Oblique = attdef.Oblique
                            ad.Thickness = attdef.Thickness
                            ad.Transparency = attdef.Transparency
                            ad.VerticalMode = attdef.VerticalMode
                            ad.Visible = attdef.Visible
                            ad.WidthFactor = attdef.WidthFactor

                            ad.CastShadows = attdef.CastShadows
                            ad.Constant = attdef.Constant
                            ad.FieldLength = attdef.FieldLength
                            ad.ForceAnnoAllVisible = attdef.ForceAnnoAllVisible
                            ad.Preset = attdef.Preset
                            ad.Prompt = attdef.Prompt
                            ad.Verifiable = attdef.Verifiable

                            ad.AdjustAlignment(db)
                        Next

                        '3. Add the missing
                        For Each attdef As AttributeDefinition In attdefs.Where(Function(n) Not dtags2.Contains(n.Tag))
                            Dim ad As New AttributeDefinition()
                            ad.SetDatabaseDefaults()
                            ad.Position = attdef.Position
                            ad.TextStyleId = attdef.TextStyleId
                            ad.TextString = attdef.TextString
                            ad.Tag = attdef.Tag
                            ad.Prompt = attdef.Prompt

                            ad.LayerId = attdef.LayerId
                            ad.Rotation = attdef.Rotation
                            ad.LinetypeId = attdef.LinetypeId
                            ad.LineWeight = attdef.LineWeight
                            ad.LinetypeScale = attdef.LinetypeScale
                            ad.Annotative = attdef.Annotative
                            ad.Color = attdef.Color
                            ad.Height = attdef.Height
                            ad.HorizontalMode = attdef.HorizontalMode
                            ad.Invisible = attdef.Invisible
                            ad.IsMirroredInX = attdef.IsMirroredInX
                            ad.IsMirroredInY = attdef.IsMirroredInY
                            ad.Justify = attdef.Justify
                            ad.LockPositionInBlock = attdef.LockPositionInBlock
                            ad.MaterialId = attdef.MaterialId
                            ad.Oblique = attdef.Oblique
                            ad.Thickness = attdef.Thickness
                            ad.Transparency = attdef.Transparency
                            ad.VerticalMode = attdef.VerticalMode
                            ad.Visible = attdef.Visible
                            ad.WidthFactor = attdef.WidthFactor

                            ad.CastShadows = attdef.CastShadows
                            ad.Constant = attdef.Constant
                            ad.FieldLength = attdef.FieldLength
                            ad.ForceAnnoAllVisible = attdef.ForceAnnoAllVisible
                            ad.Preset = attdef.Preset
                            ad.Prompt = attdef.Prompt
                            ad.Verifiable = attdef.Verifiable

                            _btr.AppendEntity(ad)
                            t.AddNewlyCreatedDBObject(ad, True)
                            ad.AdjustAlignment(db)
                        Next
                        'Sync all occurrences of the anonymous block definition
                        _btr.AttSync(directOnly, removeSuperfluous, setAttDefValues)
                    Next
                    'Update the geometry definitions of anonymous blocks derived from
                    'this dynamic block
                    btr.UpdateAnonymousBlocks()
                    t.Commit()
                End Using
            End If
        End Using
    End Sub

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
Structure KVP(Of T)
    Public Key As KeyValuePair(Of Integer, String)
    Public Value As T
End Structure
