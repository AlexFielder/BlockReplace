Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices.Core

Module ZoomExtensions
    Sub New()
    End Sub
    ' Returns the transformation matrix from the ViewTableRecord DCS to WCS
    <System.Runtime.CompilerServices.Extension> _
    Public Function EyeToWorld(view As ViewTableRecord) As Matrix3d
        Return Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target) * Matrix3d.Displacement(view.Target - Point3d.Origin) * Matrix3d.PlaneToWorld(view.ViewDirection)
    End Function

    ' Returns the transformation matrix from WCS to the ViewTableRecord DCS
    <System.Runtime.CompilerServices.Extension> _
    Public Function WorldToEye(view As ViewTableRecord) As Matrix3d
        Return view.EyeToWorld().Inverse()
    End Function

    ' Process a zoom according to the extents3d in the current viewport
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

    ' Process a zoom extents in the current viewport
    <System.Runtime.CompilerServices.Extension> _
    Public Sub ZoomExtents(ed As Editor)
        Dim db As Database = ed.Document.Database
        Dim ext As Extents3d = If(CShort(Application.GetSystemVariable("cvport")) = 1, New Extents3d(db.Pextmin, db.Pextmax), New Extents3d(db.Extmin, db.Extmax))
        ed.Zoom(ext)
    End Sub


End Module
