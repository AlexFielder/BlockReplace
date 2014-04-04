Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices.Core
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
            ' Get the block table for the current database
            Dim blockTable = DirectCast(tr.GetObject(database.BlockTableId, OpenMode.ForRead), BlockTable)

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
    ''' Returns the transformation matrix from the ViewTableRecord DCS to WCS
    ''' </summary>
    ''' <param name="view"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension> _
    Public Function EyeToWorld(view As ViewTableRecord) As Matrix3d
        Return Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) * Matrix3d.Displacement(view.Target - Point3d.Origin) * Matrix3d.PlaneToWorld(view.ViewDirection)
    End Function

    ''' <summary>
    ''' Returns the transformation matrix from WCS to the ViewTableRecord DCS
    ''' </summary>
    ''' <param name="view"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension> _
    Public Function WorldToEye(view As ViewTableRecord) As Matrix3d
        Return view.EyeToWorld().Inverse()
    End Function

    ''' <summary>
    ''' Process a zoom according to the extents3d in the current viewport
    ''' </summary>
    ''' <param name="ed"></param>
    ''' <param name="ext"></param>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension> _
    Public Sub Zoom(ed As Editor, ext As Extents3d)
        Using view As ViewTableRecord = ed.GetCurrentView()
            ext.TransformBy(view.WorldToEye())
            view.Width = ext.MaxPoint.X - ext.MinPoint.X
            view.Height = ext.MaxPoint.Y - ext.MinPoint.Y
            view.CenterPoint = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2.0, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2.0)
            ed.SetCurrentView(view)
        End Using
    End Sub

    ''' <summary>
    ''' Process a zoom extents in the current viewport
    ''' </summary>
    ''' <param name="ed"></param>
    ''' <remarks></remarks>
    <System.Runtime.CompilerServices.Extension> _
    Public Sub ZoomExtents(ed As Editor)
        Dim db As Database = ed.Document.Database
        Dim ext As Extents3d = If(CShort(Application.GetSystemVariable("cvport")) = 1, New Extents3d(db.Pextmin, db.Pextmax), New Extents3d(db.Extmin, db.Extmax))
        ed.Zoom(ext)
    End Sub
End Module
