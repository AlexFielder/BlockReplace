Imports System.Configuration
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Text.RegularExpressions.Regex
Imports System.Xml
Imports System.Xml.Serialization
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.ApplicationServices.Core
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(BlockReplace.MyCommands))> 
Namespace BlockReplace

    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class MyCommands
#Region "Constants"
        ' Type of resbuf element
        ' credit to Alexander Rivilis 
        Const RTNONE As Short = 5000
        ' No result 
        Const RTREAL As Short = 5001
        ' Real number 
        Const RTPOshort As Short = 5002
        ' 2D poshort X and Y only 
        Const RTSHORT As Short = 5003
        ' Short integer 
        Const RTANG As Short = 5004
        ' Angle 
        Const RTSTR As Short = 5005
        ' String 
        Const RTENAME As Short = 5006
        ' Entity name 
        Const RTPICKS As Short = 5007
        ' Pick set 
        Const RTORshort As Short = 5008
        ' Orientation 
        Const RT3DPOshort As Short = 5009
        ' 3D poshort - X, Y, and Z 
        Const RTLONG As Short = 5010
        ' Long integer 
        Const RTVOID As Short = 5014
        ' Blank symbol 
        Const RTLB As Short = 5016
        ' list begin 
        Const RTLE As Short = 5017
        ' list end 
        Const RTDOTE As Short = 5018
        ' dotted pair 
        Const RTNIL As Short = 5019
        ' nil 
        Const RTDXF0 As Short = 5020
        ' DXF code 0 for ads_buildlist only 
        Const RTT As Short = 5021
        ' T atom 
        Const RTRESBUF As Short = 5023
        ' resbuf 
        Const RTMODELESS As Short = 5027
        ' interrupted by modeless dialog 
        '
        ' Error return code
        Const RTNORM As Short = 5100
        ' Request succeeded 
        Const RTERROR As Short = -5001
        ' Some other error 
        Const RTCAN As Short = -5002
        ' User cancelled request -- Ctl-C 
        Const RTREJ As Short = -5003
        ' AutoCAD rejected request -- invalid 
        Const RTFAIL As Short = -5004
        ' Link failure -- Lisp probably died 
        Const RTKWORD As Short = -5005
        ' Keyword returned from getxxx() routine 
        Const RTINPUTTRUNCATED As Short = -5008
        ' Input didn't all fit in the buffer 
#End Region

        Dim DoneAddingPoints As Boolean
        Dim KeepAddingAttRefs As Boolean
        Dim KeepAddingCrossingLines As Boolean
        Dim UsedOnImgUrl As String
        Dim SelEntId As ObjectId
        Dim SelEntIds As ObjectId()
        Dim tmpblkname As String = String.Empty
        Dim tmpm As New mappings
        Dim d As New Drawings
        Dim drawing As DrawingsDrawing
        Dim Snapshots As New Snapshots ' new snapshot details parent collection
        Dim snDetails As SnapshotsSnapshotDetails ' new snapshots collection
        Dim snapshot As SnapshotsSnapshotDetailsSnapshot ' new snapshot detail
        ''' <summary>
        ''' A list of Revision
        ''' </summary>
        Public Shared revisions As New List(Of Revision)()

        Private Property DrawingID As String

        ''' <summary>
        ''' Allows the user to manually select a block for replacement.
        ''' </summary>
        ''' <remarks></remarks>
        <CommandMethod("UpdateAttributes")> _
        Public Sub UpdateAttributes()
            Dim prmptSelOpts As New PromptSelectionOptions()
            Dim prmptSelRes As PromptSelectionResult
            Dim rbfResult As ResultBuffer

            prmptSelRes = Active.Editor.GetSelection(prmptSelOpts)
            If prmptSelRes.Status <> PromptStatus.OK Then
                'the user didn't select anything.
                Exit Sub
            End If
            Dim selSet As SelectionSet
            selSet = prmptSelRes.Value
            rbfResult = New ResultBuffer( _
                        New TypedValue(LispDataType.SelectionSet, selSet))
            UpdateAtts(rbfResult)
            rbfResult.Dispose()
        End Sub

        ''' <summary>
        ''' Allows the user to automate changing of blocks from the command line.
        ''' Works particularly well with Scriptpro.
        ''' </summary>
        ''' <param name="args">the selectionset required to grab the relevant objectid</param>
        ''' <remarks></remarks>
        <LispFunction("UpdateAtts")> _
        Public Sub UpdateAtts(args As ResultBuffer)
            If args Is Nothing Then
                Throw New ArgumentException("Requires one argument")
            End If
            Dim values As TypedValue() = args.AsArray
            If values.Length <> 1 Then
                Throw New ArgumentException("Wrong number of arguments")
            End If
            If values(0).TypeCode <> CInt(LispDataType.SelectionSet) Then
                Throw New ArgumentException("Bad argument type - requires a selection set")
            End If
            Dim ss As SelectionSet = DirectCast(values(0).Value, SelectionSet)
            'Dim selEntIDs As ObjectId()
            Dim asmpath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim tmpserializer As New XmlSerializer(GetType(mappings))
            Dim tmpfs As New FileStream(asmpath + "\Resources\Mappings.xml", FileMode.Open)
            Dim tmpreader As XmlReader = XmlReader.Create(tmpfs)
            'new these objects here.
            snDetails = New SnapshotsSnapshotDetails
            drawing = New DrawingsDrawing
            'then continue
            Dim originalfilename As String = Active.Database.Filename
            Dim g As Guid = Guid.NewGuid()
            DrawingID = g.ToString()
            drawing.DrawingID = DrawingID
            Dim fn As String = "C:\temp\" & Path.GetFileNameWithoutExtension(originalfilename) & "_Before.png"
            Dim ext As Extents3d = If(CShort(Application.GetSystemVariable("cvport")) = 1, _
                                      New Extents3d(Active.Database.Pextmin, Active.Database.Pextmax), _
                                      New Extents3d(Active.Database.Extmin, Active.Database.Extmax))
            drawing.BeforeImgURL = ScreenShotToFile(ext.MinPoint, ext.MaxPoint, fn)
            
            tmpm = CType(tmpserializer.Deserialize(tmpreader), mappings)
            tmpfs.Close()
            If ss.Count > 1 Then
                ForEach(Of BlockReference)(Active.Database, AddressOf ProcessBlockReferences)
            Else
                SelEntIds = ss.GetObjectIds() 'currently this doesn't play nice if there happens to be a second attributed block in the drawing!
                SelEntId = SelEntIds(0)
            End If
            'we can now choose to drop this part of the tool- if the before & after snapshots are good enough!
            'CreateSnapshotsForAreasOfInterest(ss, asmpath)
            tmpfs.Dispose()
            Dim selEntID2 As ObjectId '=  Active.editor.GetEntity("Select New block:").ObjectId
            Using doclock As DocumentLock = Active.Document.LockDocument 'lock the document whilst we edit it.
                Using myTrans As Transaction = Active.Database.TransactionManager.StartTransaction
                    Dim myBrefA As BlockReference = selEntID.GetObject(OpenMode.ForWrite)
                    Dim myBrefB As BlockReference '= selEntID2.GetObject(OpenMode.ForWrite)
                    Dim myAttsA As AttributeCollection = myBrefA.AttributeCollection
                    Dim myAttsB As AttributeCollection '= myBrefB.AttributeCollection
                    Dim blockNameA As String = ""
                    Dim blockNameB As String = ""
                    Dim serializerm As New XmlSerializer(GetType(mappings))
                    Dim serializerf As New XmlSerializer(GetType(frames))
                    Dim serializerd As New XmlSerializer(GetType(Drawings))

                    Dim fsbr As New FileStream(asmpath + "\Resources\BlockReplace.xml", FileMode.Open)
                    Dim fsm As New FileStream(asmpath + "\Resources\Mappings.xml", FileMode.Open)
                    Dim readerBR As XmlReader = XmlReader.Create(fsbr)
                    Dim readerM As XmlReader = XmlReader.Create(fsm)
                    Dim f As frames
                    Dim m As mappings
                    drawing.oldname = Application.GetSystemVariable("DWGNAME") 'Active.Database.Filename
                    drawing.oldpath = Application.GetSystemVariable("DWGPREFIX")

                    f = CType(serializerf.Deserialize(readerBR), frames)
                    m = CType(serializerm.Deserialize(readerM), mappings)
                    fsbr.Close()
                    fsm.Close()
                    If myBrefA.Name.StartsWith("*") Then
                        Dim myBTR As BlockTableRecord = myBrefA.DynamicBlockTableRecord.GetObject(OpenMode.ForRead)
                        blockNameA = myBTR.Name
                    Else
                        blockNameA = myBrefA.Name
                    End If
                    If Not blockNameA Like "*5.2(block)" Then

                        Dim ob As mappingsOldblock = (From s In m.oldblock
                                                      Where s.name = blockNameA
                                                      Select s).SingleOrDefault()
                        Dim myBT As BlockTable = Active.Database.BlockTableId.GetObject(OpenMode.ForWrite)
                        If myBT.Has(ob.newname) = False Then
                            'insert DWG file as a Block
                            Dim myDWG As New IO.FileInfo(ob.path)
                            If myDWG.Exists = False Then
                                MsgBox("The file " & myDWG.FullName & " does not exist.")
                                Exit Sub
                            End If
                            'Create a blank Database
                            Dim dwgDB As New Database(False, True)
                            'Read a DWG file into the blank database
                            dwgDB.ReadDwgFile(myDWG.FullName, FileOpenMode.OpenForReadAndAllShare, True, "")
                            'insert the dwg file into the current file's block table
                            Active.Database.Insert(ob.newname, dwgDB, True)
                            'close/dispose of the previously blank database.
                            dwgDB.Dispose()
                        End If
                        Dim blockPos As Point3d
                        If blockNameA = "A3BORD" Then
                            blockPos = New Point3d(myBrefA.Position.X - 512.145, myBrefA.Position.Y, myBrefA.Position.Z)
                        ElseIf blockNameA = "A2 BORDER" Then
                            blockPos = New Point3d(myBrefA.Position.X + 299.372, myBrefA.Position.Y + 160.3367, myBrefA.Position.Z)
                        Else
                            blockPos = myBrefA.Position
                        End If
                        selEntID2 = InsertBlock(Active.Database, _
                                                myBrefA.BlockName, _
                                                blockPos, _
                                                ob.newname, _
                                                myBrefA.ScaleFactors.X, _
                                                myBrefA.ScaleFactors.Y, _
                                                myBrefA.ScaleFactors.Z)
                        myBrefB = selEntID2.GetObject(OpenMode.ForWrite)
                        blockNameB = myBrefB.Name
                        myAttsB = myBrefB.AttributeCollection

                        For Each att In ob.attributes
                            For Each myAttID As ObjectId In myAttsA
                                Dim myAtt As AttributeReference = myAttID.GetObject(OpenMode.ForRead)
                                If myAtt.Tag.ToUpper = att.name.ToUpper Then
                                    For Each myAttBID As ObjectId In myAttsB
                                        Dim myAttB As AttributeReference = myAttBID.GetObject(OpenMode.ForWrite)
                                        If myAttB.Tag.ToUpper = att.newname.ToUpper Then
                                            myAttB.TextString = getCorrectedDrawingNumber(Active.Document, myAtt, myAttB.Tag, blockNameA)
                                            'myAttB.TextString = myAtt.TextString
                                        End If
                                    Next
                                End If
                            Next
                        Next
                    Else ' we have a 5.2 version frame already & should update it to our latest version.
                        'this doesn't work for whatever reason. reverted to original code for now.
                        'myBrefB = myTrans.GetObject(ReplaceBlock(myBrefA.ObjectId), OpenMode.ForWrite)
                        myBrefB = myTrans.GetObject(myBrefA.ObjectId, OpenMode.ForRead)
                        blockNameB = myBrefB.Name
                        myAttsB = myBrefB.AttributeCollection
                    End If
                    'open the blockreplace.xml file to grab dumb text from any of the frames
                    'and populate our new drawing frame
                    Dim WeKnowAboutThisBorder = (From a In f.DrawingFrame
                                                 Where a.name = blockNameA
                                                 Select a).SingleOrDefault()
                    If WeKnowAboutThisBorder IsNot Nothing Then
                        Dim tmppntcoll As Point3dCollection = New Point3dCollection
                        Dim sb = (From s In f.DrawingFrame
                                  Where s.name = blockNameA
                                  Let sbb = s.SearchBoxBounds
                                  Select sbb).SingleOrDefault()
                        For Each att As framesDrawingFrameSearchBoxBoundsAttref In sb.Attrefs
                            tmppntcoll = New Point3dCollection
                            If att.name = "USED_ON" Then
                                Dim tmpstrlist As List(Of String) = New List(Of String)
                                For Each Pnt3d In att.point3d
                                    tmppntcoll.Add(New Point3d(Pnt3d.X + myBrefB.Position.X, Pnt3d.Y + myBrefB.Position.Y, 0))
                                Next
                                tmpstrlist = CollectTextListFromArea(tmppntcoll, att.name)
                                If tmpstrlist.Count > 0 Then
                                    For i As Integer = 0 To tmpstrlist.Count - 1
                                        Dim tmpint As Integer = i + 1
                                        Dim pad As Char = "0"c
                                        Dim str As String = "USED ON " + CStr(tmpint).PadLeft(2, pad) + " - PREFIX"
                                        Dim attr As AttributeReference = (From tmpatt As ObjectId In myAttsB
                                                                          Let attref As AttributeReference = tmpatt.GetObject(OpenMode.ForWrite)
                                                                          Where attref.Tag = str
                                                                          Select attref).SingleOrDefault()
                                        attr.TextString = tmpstrlist.Item(i)
                                        PadSnapshotsForAreasOfInterest(att, attr, str, True)
                                    Next
                                    tmpstrlist = Nothing
                                    tmppntcoll = Nothing
                                End If
                            Else
                                Dim tmpstr As String = String.Empty
                                For Each Pnt3d In att.point3d
                                    tmppntcoll.Add(New Point3d(Pnt3d.X + myBrefB.Position.X, Pnt3d.Y + myBrefB.Position.Y, 0))
                                Next
                                Dim linecount As Integer = 0
                                tmpstr = CollectTextFromArea(tmppntcoll:=tmppntcoll, locName:=att.name, reverseSort:=False, lineCount:=linecount)
                                Dim attr As AttributeReference = (From tmpatt As ObjectId In myAttsB
                                                                  Let attref As AttributeReference = tmpatt.GetObject(OpenMode.ForWrite)
                                                                  Where attref.Tag = att.AttributeName
                                                                  Select attref).SingleOrDefault()
                                If Not attr = Nothing Then
                                    If attr.TextString = "" Then
                                        attr.TextString = tmpstr
                                    End If
                                    If attr.Tag = "TOLERANCE" And Not attr.TextString = "" Then 'always going to be this one that overlaps now.
                                        If linecount > 2 Then '3 or more lines
                                            Dim pnt As Point3d = attr.AlignmentPoint
                                            attr.AlignmentPoint = New Point3d(pnt.X, pnt.Y + (attr.Height * 2), pnt.Z)
                                        End If
                                        attr.UpdateMTextAttribute()
                                    End If
                                    PadSnapshotsForAreasOfInterest(att, attr, tmpstr, False)
                                End If
                                End If
                        Next
                    End If

                    'move text notes out of the way but only for non-Item List drawings.
                    If Not myBrefB.Name Like "*IL*" Then
                        GroupAndMoveTextNotes(ext, myBrefB.ObjectId)
                    End If
                    'then delete extra lines, text, mtext and blocks which are in the original border.
                    Dim sblines = (From s In f.DrawingFrame
                              Where s.name = blockNameA
                              Let sbb = s.SearchBoxBounds
                              Select sbb).SingleOrDefault()
                    Dim pnts As Point3dCollection = New Point3dCollection()
                    For Each pnt In sblines.Points
                        pnts.Add(New Point3d(pnt.X + myBrefB.Position.X, pnt.Y + myBrefB.Position.Y, 0))
                    Next
                    DeleteMySelection(pnts, False)

                    'then do the same for our new border
                    sblines = (From s In f.DrawingFrame
                              Where s.name = blockNameB
                              Let sbb = s.SearchBoxBounds
                              Select sbb).SingleOrDefault()
                    pnts = New Point3dCollection()
                    For Each pnt In sblines.Points
                        pnts.Add(New Point3d(pnt.X + myBrefB.Position.X, pnt.Y + myBrefB.Position.Y, 0))
                    Next

                    DeleteMySelection(pnts, False)

                    Active.WriteMessage(vbCrLf & "Done deleting old lines!" & vbCrLf)
                    Dim lines = (From ln In f.DrawingFrame
                                 Where ln.name = blockNameB
                                 Let sbl = ln.Lines
                                 Select sbl).SingleOrDefault()

                    For Each Xline As framesDrawingFrameLine In lines
                        Dim tmpattr As AttributeReference = (From a As ObjectId In myAttsB
                                                             Let attr As AttributeReference = a.GetObject(OpenMode.ForRead)
                                                             Where attr.Tag = Xline.name
                                                             Select attr).SingleOrDefault()
                        If Not tmpattr = Nothing Then
                            If tmpattr.TextString = "" Then
                                Dim linestart As Point3d = New Point3d(Xline.LineStart.X + myBrefB.Position.X, Xline.LineStart.Y + myBrefB.Position.Y, 0)
                                Dim lineend As Point3d = New Point3d(Xline.LineEnd.X + myBrefB.Position.X, Xline.LineEnd.Y + myBrefB.Position.Y, 0)
                                Dim ln = New Line(linestart, lineend)
                                Dim btr = DirectCast(myTrans.GetObject(Active.Database.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                                btr.AppendEntity(ln)
                                ln.Layer = Xline.Layer
                                myTrans.AddNewlyCreatedDBObject(ln, True)
                                linestart = Nothing
                                lineend = Nothing
                            End If
                        End If

                    Next
                    If Not myBrefA.ObjectId = myBrefB.ObjectId Then
                        myBrefA.Erase()
                    End If
                    d.Drawing.Add(drawing)
                    Snapshots.SnapshotDetails.Add(snDetails)
                    Dim fs As FileStream
                    If Not File.Exists("C:\Temp\Drawings.xml") Then
                        fs = New FileStream("C:\Temp\Drawings.xml", FileMode.Create)
                        Dim writer As New XmlTextWriter(fs, Encoding.Unicode)
                        writer.Formatting = Formatting.Indented
                        serializerd.Serialize(writer, d) ' so we include the xml header..?
                        writer.Close()
                    Else
                        'open the existing xml file
                        Dim serializernew As New XmlSerializer(GetType(Drawings))
                        Dim fsdrawing As New FileStream("C:\Temp\Drawings.xml", FileMode.Open)
                        Dim readerDrawing As XmlReader = XmlReader.Create(fsdrawing)
                        Dim r As Drawings
                        r = CType(serializernew.Deserialize(readerDrawing), Drawings)
                        fsdrawing.Close()
                        readerDrawing = Nothing
                        'overwrite the contents with our updated data.
                        'to cope with re-runs we need to add this:
                        Dim dwg = (From a As DrawingsDrawing In r.Drawing
                                   Where Path.GetFileNameWithoutExtension(a.oldname) = _
                                   Path.GetFileNameWithoutExtension(originalfilename)
                                   Select a).FirstOrDefault()
                        If dwg IsNot Nothing Then
                            'if the file already exists in drawings.xml then we don't want to override the DrawingID as this would break
                            'the link to the snapshots dataset; instead we swap in the new/updated data:
                            dwg.BeforeImgURL = drawing.BeforeImgURL
                            dwg.name = drawing.name
                            dwg.oldname = drawing.oldname
                            dwg.oldpath = drawing.oldpath
                            dwg.path = drawing.path
                            dwg.revision = drawing.revision
                            dwg.revisiondatestr = drawing.revisiondatestr
                        Else
                            'otherwise just add the new drawing to the list:
                            r.Drawing.Add(drawing)
                        End If
                        fs = New FileStream("C:\Temp\Drawings.xml", FileMode.Create)
                        Dim writer As New XmlTextWriter(fs, Encoding.Unicode)
                        writer.Formatting = Formatting.Indented
                        serializerd.Serialize(writer, r) 'so we only include our new drawing.
                        writer.Close()
                    End If
                    Dim serializers As New XmlSerializer(GetType(Snapshots))
                    If Not File.Exists("C:\Temp\Snapshots.xml") Then
                        fs = New FileStream("C:\Temp\Snapshots.xml", FileMode.Create)
                        Dim writer As New XmlTextWriter(fs, Encoding.Unicode)
                        writer.Formatting = Formatting.Indented
                        serializers.Serialize(writer, Snapshots) ' so we include the xml header..?
                        writer.Close()
                    Else
                        'open the existing xml file
                        Dim serializernew As New XmlSerializer(GetType(Snapshots))
                        Dim fsdrawing As New FileStream("C:\Temp\Snapshots.xml", FileMode.Open)
                        Dim readerDrawing As XmlReader = XmlReader.Create(fsdrawing)
                        Dim sn As Snapshots
                        sn = CType(serializernew.Deserialize(readerDrawing), Snapshots)
                        fsdrawing.Close()
                        readerDrawing = Nothing
                        'overwrite the contents with our updated data.
                        sn.SnapshotDetails.Add(snDetails)
                        fs = New FileStream("C:\Temp\Snapshots.xml", FileMode.Create)
                        Dim writer As New XmlTextWriter(fs, Encoding.Unicode)
                        writer.Formatting = Formatting.Indented
                        serializers.Serialize(writer, sn) 'so we only include our new drawing.
                        writer.Close()
                    End If
                    myTrans.Commit()
                End Using
            End Using
            'cleanup some things between drawings or FAIL.
            drawing = Nothing
            snDetails = Nothing
        End Sub

        ''' <summary>
        ''' Function that allows us to filter for the blockreference.ObjectId we want
        ''' </summary>
        ''' <param name="blkref"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ProcessBlockReferences(blkref As BlockReference) As Action(Of BlockReference)
            Dim tr As Transaction = Active.Database.TransactionManager.StartTransaction
            Try
                If SelEntId = ObjectId.Null Then
                    If blkref.Name.StartsWith("*") Then
                        tmpblkname = _
                            DirectCast(blkref.DynamicBlockTableRecord.GetObject(OpenMode.ForRead),  _
                                BlockTableRecord).Name
                        Dim ob As mappingsOldblock = (From s In tmpm.oldblock
                                                  Where s.name = tmpblkname
                                                  Select s).SingleOrDefault()
                        If Not ob Is Nothing Then
                            SelEntId = blkref.ObjectId
                        Else
                            SelEntId = ObjectId.Null
                        End If
                    Else
                        If Not blkref.Name Like "*5.2(block)" Then 'go look and see if we know about this block already.
                            tmpblkname = blkref.Name
                            Dim ob As mappingsOldblock = (From s In tmpm.oldblock
                                                          Where s.name = tmpblkname
                                                          Select s).SingleOrDefault()
                            If Not ob Is Nothing Then
                                ob = Nothing
                                SelEntId = blkref.ObjectId
                                'Exit Function
                            Else
                                SelEntId = ObjectId.Null
                            End If
                        Else 'we don't need to do anything to this block, just capture it's id.
                            SelEntId = blkref.ObjectId
                        End If
                    End If
                End If
            Catch ex As Exception
                Active.WriteMessage("The Error was: " & ex.Message)
            Finally
                tr.Commit()
                tr.Dispose()
            End Try

        End Function

        ''' <summary>
        ''' Creates snapshots of areas based on the selectionset passed to it.
        ''' </summary>
        ''' <param name="ss"></param>
        ''' <param name="asmpath"></param>
        ''' <remarks></remarks>
        Private Sub CreateSnapshotsForAreasOfInterest(ss As SelectionSet, asmpath As String)
            Dim fn As String = String.Empty
            Using tmptrans As Transaction = Active.Database.TransactionManager.StartTransaction
                If Not SelEntId = Nothing Then
                    'capture a screenshot of each existing area where text might be located:
                    drawing.oldname = Application.GetSystemVariable("DWGNAME") 'Active.Database.Filename
                    drawing.oldpath = Application.GetSystemVariable("DWGPREFIX")
                    Dim serializerf As New XmlSerializer(GetType(frames))
                    Dim fsbr As New FileStream(asmpath + "\Resources\BlockReplace.xml", FileMode.Open)
                    Dim readerBR As XmlReader = XmlReader.Create(fsbr)
                    Dim f As frames = CType(serializerf.Deserialize(readerBR), frames)
                    fsbr.Close()
                    Dim brefa As BlockReference = SelEntId.GetObject(OpenMode.ForRead)
                    tmpblkname = brefa.Name
                    Dim WeKnowAboutThisBorder = (From a In f.DrawingFrame
                                                     Where a.name = tmpblkname
                                                     Select a).SingleOrDefault()
                    If WeKnowAboutThisBorder IsNot Nothing Then
                        Dim sb = (From s In f.DrawingFrame
                              Where s.name = tmpblkname
                              Let sbb = s.SearchBoxBounds
                              Select sbb).SingleOrDefault()

                        For Each att As framesDrawingFrameSearchBoxBoundsAttref In sb.Attrefs
                            snapshot = New SnapshotsSnapshotDetailsSnapshot
                            Dim tmppntcoll As Point3dCollection = New Point3dCollection
                            For Each Pnt3d In att.point3d
                                tmppntcoll.Add(New Point3d(Pnt3d.X + brefa.Position.X, Pnt3d.Y + brefa.Position.Y, 0))
                            Next
                            Dim pnt1 As Point3d
                            Dim pnt2 As Point3d
                            If brefa.Name = "A3BORD" Then 'don't add the block position to the screenshot points
                                pnt1 = New Point3d(att.point3d.Item(0).X, att.point3d.Item(0).Y, att.point3d.Item(0).Z)
                                pnt2 = New Point3d(att.point3d.Item(2).X, att.point3d.Item(2).Y, att.point3d.Item(2).Z)
                            Else
                                pnt1 = New Point3d(att.point3d.Item(0).X + brefa.Position.X, att.point3d.Item(0).Y + brefa.Position.Y, att.point3d.Item(0).Z)
                                pnt2 = New Point3d(att.point3d.Item(2).X + brefa.Position.X, att.point3d.Item(2).Y + brefa.Position.Y, att.point3d.Item(2).Z)
                            End If

                            fn = "C:\temp\" & Path.GetFileNameWithoutExtension(Active.Database.Filename) & "_" & att.name & ".png"
                            snapshot.name = att.name
                            'snapshot.ImgUrl = CaptureSnapShot(pnt1, pnt2, fn)
                            snapshot.ImgUrl = ScreenShotToFile(pnt1, pnt2, fn)
                            If att.name = "USED_ON" Then
                                UsedOnImgUrl = snapshot.ImgUrl
                            End If
                            Dim pntmin As New SnapshotsSnapshotDetailsSnapshotPoint3d With {.name = att.name, .X = pnt1.X, .Y = pnt1.Y, .Z = 0}
                            Dim pntmax As New SnapshotsSnapshotDetailsSnapshotPoint3d With {.name = att.name, .X = pnt2.X, .Y = pnt2.Y, .Z = 0}
                            snapshot.capturedarea.Add(pntmin)
                            snapshot.capturedarea.Add(pntmax)
                            snDetails.DrawingID = DrawingID
                            snDetails.snapshot.Add(snapshot)
                        Next
                        Active.WriteMessage(vbCrLf & "Done Capturing Area Screenshots" & vbCrLf)
                    Else
                        Active.WriteMessage("This isn't a drawing we have had prior knowledge of!" & vbCrLf & "Suggest you run the GetPointsFromUser tool and try again!")
                        Exit Sub
                    End If
                End If
            End Using
        End Sub

        ''' <summary>
        ''' Adds extra information to the snapshots we previously created.
        ''' </summary>
        ''' <param name="att"></param>
        ''' <param name="attr"></param>
        ''' <param name="str"></param>
        ''' <param name="isUsedOnAttr"></param>
        ''' <remarks></remarks>
        Private Sub PadSnapshotsForAreasOfInterest(att As framesDrawingFrameSearchBoxBoundsAttref, attr As AttributeReference, str As String, isUsedOnAttr As Boolean)
            If snapshot Is Nothing Then snapshot = New SnapshotsSnapshotDetailsSnapshot
            snapshot = (From sn As SnapshotsSnapshotDetailsSnapshot In snDetails.snapshot
                        Where sn.name = att.name
                        Select sn).SingleOrDefault()
            If snapshot Is Nothing Then
                snapshot = New SnapshotsSnapshotDetailsSnapshot
                If isUsedOnAttr Then
                    snapshot.ImgUrl = UsedOnImgUrl
                    snapshot.name = str
                    If snapshot.tag Is Nothing Or snapshot.tag = "" Then
                        snapshot.tag = str
                    End If
                Else
                    snapshot.name = str
                    If snapshot.tag Is Nothing Or snapshot.tag = "" Then
                        snapshot.tag = attr.Tag
                    End If
                End If

                snapshot.objectIdAsString = attr.ObjectId.ToString()
                snapshot.textstring = attr.TextString
                snDetails.snapshot.Add(snapshot)
            Else
                snapshot.objectIdAsString = attr.ObjectId.ToString()
                snapshot.tag = attr.Tag
                snapshot.textstring = attr.TextString
            End If
            'snapshot = (From sn As SnapshotsSnapshotDetailsSnapshot In snDetails.snapshot
            '            Where sn.name = att.name
            '            Select sn).SingleOrDefault()
            'If snapshot Is Nothing Or Not snapshot.tag = "" Then 'assume we already have a USED ON 01 - PREFIX
            '    snapshot = New SnapshotsSnapshotDetailsSnapshot
            '    snapshot.ImgUrl = UsedOnImgUrl
            '    snapshot.name = tmpstr
            '    snapshot.tag = attr.Tag
            '    snapshot.objectIdAsString = attr.ObjectId.ToString()
            '    snapshot.textstring = attr.TextString
            '    snDetails.snapshot.Add(snapshot)
            'Else
            '    snapshot.objectIdAsString = attr.ObjectId.ToString()
            '    snapshot.tag = attr.Tag
            '    snapshot.textstring = attr.TextString
            '    'snDetails.snapshot.Add(snapshot)
            'End If
        End Sub

        ''' <summary>
        ''' Creates a screenshot of each area we are interested in.
        ''' </summary>
        ''' <param name="pt1">Point3d origin</param>
        ''' <param name="pt2">Point3d extents</param>
        ''' <param name="filename">Expected File Output</param>
        ''' <returns>Returns the filename output for inclusiong in reports.xml</returns>
        ''' <remarks></remarks>
        Private Shared Function ScreenShotToFile(pt1 As Point3d, pt2 As Point3d, filename As String) As String
            'Dim ed As Editor = Active.Editor
            Dim cvport = CShort(Application.GetSystemVariable("CVPORT"))
            Dim img As System.Drawing.Image
            Dim view As ViewTableRecord = Active.Editor.GetCurrentView()
            'start transaction
            Using trans As Transaction = Active.Database.TransactionManager.StartTransaction()
                'get the entity' extends
                'configure the new current view
                view.Width = pt2.X - pt1.X
                view.Height = pt2.Y - pt1.Y
                view.CenterPoint = New Point2d( _
                  pt1.X + (view.Width / 2), _
                  pt1.Y + (view.Height / 2))
                'update the view
                Active.Editor.SetCurrentView(view)
                trans.Commit()
            End Using
            Using tr As Transaction = Active.document.TransactionManager.StartTransaction()
                Dim vst As GraphicsInterface.VisualStyleType = GraphicsInterface.VisualStyleType.Basic
                Dim gsm As GraphicsSystem.Manager = Active.document.GraphicsManager
                Using gfxview As GraphicsSystem.View = New GraphicsSystem.View
                    gsm.SetViewFromViewport(gfxview, cvport)
                    gfxview.VisualStyle = New GraphicsInterface.VisualStyle(vst)
                    Using dev As GraphicsSystem.Device = gsm.CreateAutoCADOffScreenDevice()
                        Dim s As Size = New Size(gsm.DeviceIndependentDisplaySize.Width, gsm.DeviceIndependentDisplaySize.Height)
                        Dim sz As New Size(Math.Abs(pt1.X - pt2.X), Math.Abs(pt1.Y - pt2.Y))
                        dev.OnSize(s)
                        dev.DeviceRenderType = GraphicsSystem.RendererType.Default
                        dev.BackgroundColor = Color.White
                        dev.Add(gfxview)
                        dev.Update()
                        Using model As GraphicsSystem.Model = gsm.CreateAutoCADModel
                            Using trans As Transaction = Active.document.TransactionManager.StartTransaction()
                                Dim btr As BlockTableRecord = trans.GetObject(Active.Database.CurrentSpaceId, OpenMode.ForRead)
                                gfxview.Add(btr, model)
                                trans.Commit()
                            End Using
                            Dim extents As Extents2d = gfxview.ViewportExtents

                            Dim min_pt As New Point2d(extents.MinPoint.X * s.Width, extents.MinPoint.Y * s.Height)
                            Dim max_pt As New Point2d(extents.MaxPoint.X * s.Width, extents.MaxPoint.Y * s.Height)

                            Dim view_rect As New System.Drawing.Rectangle(CInt(min_pt.X), _
                                                                          CInt(min_pt.Y), _
                                                                          CInt(System.Math.Abs(max_pt.X - min_pt.X)), _
                                                                          CInt(System.Math.Abs(max_pt.Y - min_pt.Y)))
                            img = gfxview.GetSnapshot(view_rect)

                            Dim processed As Bitmap = img
                            If filename IsNot Nothing AndAlso filename <> "" Then
                                processed.Save(filename, ImageFormat.Png)

                                Active.WriteMessage(vbCrLf & "Image captured and saved to ""{0}"".", filename)
                            End If
                            'cleanup
                            processed.Dispose()
                            gfxview.EraseAll()
                            dev.Erase(gfxview)
                        End Using
                    End Using
                End Using
                tr.Commit()
            End Using
            Return filename
        End Function

        ''' <summary>
        ''' Allows the user to add unknown borders to the BlockReplace.xml file.
        ''' </summary>
        ''' <remarks></remarks>
        <CommandMethod("ADNIntersectsWith")> _
        Public Sub ADNIntersectsWith()
            Using tr As Transaction = Active.Database.TransactionManager.StartTransaction()
                Dim peo As PromptEntityOptions = New PromptEntityOptions(vbCrLf & "Select a blockreference >>")
                peo.SetRejectMessage("Only Blockreferences allowed!")
                peo.AddAllowedClass(GetType(BlockReference), False)

                Dim res As PromptEntityResult = Active.Editor.GetEntity(peo)
                If Not res.Status = PromptStatus.OK Then Return
                Dim bref As BlockReference = DirectCast(tr.GetObject(res.ObjectId, OpenMode.ForRead), BlockReference)
                If bref = Nothing Then Return

                peo = New PromptEntityOptions(vbCrLf & "Select another entity >>")
                peo.SetRejectMessage("Only Mtext allowed!")
                peo.AddAllowedClass(GetType(MText), False)
                res = Active.Editor.GetEntity(peo)
                If Not res.Status = PromptStatus.OK Then Return
                Dim ent2 As Entity = DirectCast(tr.GetObject(res.ObjectId, OpenMode.ForRead), Entity)
                Dim obIdList As New List(Of ObjectId)
                Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                'get only the objectid of entities we intersect with.
                For Each id As ObjectId In btr
                    If Not id.ObjectClass.DxfName = "ATTDEF" Then
                        Dim points As New Point3dCollection
                        Dim tmpentity As Entity = DirectCast(id.GetObject(OpenMode.ForRead), Entity)
                        ent2.IntersectWith(tmpentity, Intersect.OnBothOperands, points, IntPtr.Zero, IntPtr.Zero)
                        If points.Count > 0 Then ' we have an overlap?
                            obIdList.Add(tmpEntity.ObjectId)
                        End If
                    End If
                Next
                DetermineOverlap(ent2.ObjectId, obIdList)
                'DetermineOverlap(ent2.ObjectId, ent1.ObjectId)

                tr.Commit()
            End Using
        End Sub

        ''' <summary>
        ''' This is the automatic version now that testing is complete.
        ''' </summary>
        ''' <param name="txtid"></param>
        ''' <param name="blkid"></param>
        ''' <remarks></remarks>
        Public Sub ADNIntersectsWith(txtid As ObjectId, blkid As ObjectId)
            Using tr As Transaction = Active.Database.TransactionManager.StartTransaction()
                Dim bref As BlockReference = DirectCast(tr.GetObject(blkid, OpenMode.ForRead), BlockReference)
                If bref = Nothing Then Return
                Dim ent2 As Entity = DirectCast(tr.GetObject(txtid, OpenMode.ForRead), Entity)
                Dim obIdList As New List(Of ObjectId)
                Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bref.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                'get only the objectid of entities we intersect with.
                For Each id As ObjectId In btr
                    If Not id.ObjectClass.DxfName = "ATTDEF" Then
                        Dim points As New Point3dCollection
                        Dim tmpentity As Entity = DirectCast(id.GetObject(OpenMode.ForRead), Entity)
                        ent2.IntersectWith(tmpentity, Intersect.OnBothOperands, points, IntPtr.Zero, IntPtr.Zero)
                        If points.Count > 0 Then ' we have an overlap?
                            obIdList.Add(tmpentity.ObjectId)
                        End If
                    End If
                Next
                DetermineOverlap(ent2.ObjectId, obIdList)
                'DetermineOverlap(ent2.ObjectId, ent1.ObjectId)

                tr.Commit()
            End Using
        End Sub

        ''' <summary>
        ''' This is the "multiple" instance of this method. It ultimately calls the singular instance below.
        ''' </summary>
        ''' <param name="objectid"></param>
        ''' <param name="obids"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function DetermineOverlap(objectid As ObjectId, obids As List(Of ObjectId)) As Point3d
            Dim tr As Transaction = Active.Database.TransactionManager.StartTransaction
            Try
                Dim mtxt As MText = DirectCast(objectid.GetObject(OpenMode.ForWrite), MText)
                Dim ext As Extents3d = mtxt.GeometricExtents
                Dim pnt As Point3d
                'Dim pmtSelRes As PromptSelectionResult = Nothing
                'Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(DxfCode.Start, "*")}
                'Dim selFilter As New SelectionFilter(acTypValAr)
                'pmtSelRes = Active.Editor.SelectCrossingWindow(ext.MinPoint, ext.MaxPoint, selFilter)
                'If pmtSelRes.Status = PromptStatus.OK Then ' we found something.
                For Each ob As ObjectId In obids
                    Dim tmpEntity As Entity = DirectCast(ob.GetObject(OpenMode.ForRead), Entity)
                    Dim points As New Point3dCollection
                    Dim tmppnts As New Point3dCollection
                    mtxt.IntersectWith(tmpEntity, Intersect.OnBothOperands, tmppnts, IntPtr.Zero, IntPtr.Zero)
                    If tmppnts.Count > 0 Then 'we know we have an intersection this time.
                        tmppnts = Nothing
                        'if we new "points" here it shouldn't persist through our recursion?
                        points = New Point3dCollection
                    End If
                    mtxt.IntersectWith(tmpEntity, Intersect.OnBothOperands, points, IntPtr.Zero, IntPtr.Zero)
                    If points.Count > 0 Then ' we have an overlap?
                        Dim tmppnt As Point3d = mtxt.Location
                        pnt = points.Item(0)
                        mtxt.Location = GetBestDirectionForMove(tmppnt, pnt) 'New Point3d(tmppnt.X, tmppnt.Y + 1.0, tmppnt.Z)
                        mtxt.Location = DetermineOverlap(mtxt.ObjectId, ob)
                        pnt = mtxt.Location
                    Else 'no overlap?
                        pnt = mtxt.Location
                        'Return pnt
                    End If
                Next
                Return pnt
                'Else
                'pnt = mtxt.Location
                'Return pnt
                'End If
            Catch ex As Exception
                Active.WriteMessage(vbCrLf & "The error was: " & ex.Message)
            Finally
                tr.Commit()
                tr.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' This is the singular instance of this method.
        ''' </summary>
        ''' <param name="objectId">ObjectId of the object we're checking.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function DetermineOverlap(objectId As ObjectId, tmpentId As ObjectId) As Point3d
            Dim tr = Active.Database.TransactionManager.StartTransaction()
            Try
                Dim mtxt As MText = DirectCast(objectId.GetObject(OpenMode.ForWrite), MText)
                Dim tmpent As Entity = DirectCast(tmpentId.GetObject(OpenMode.ForRead), Entity)
                Dim ext As Extents3d = mtxt.GeometricExtents
                Dim pnt As Point3d
                Dim pmtSelRes As PromptSelectionResult = Nothing
                Dim acTypValAr As TypedValue() = New TypedValue() {New TypedValue(DxfCode.Start, "*")}
                Dim selFilter As New SelectionFilter(acTypValAr)
                pmtSelRes = Active.Editor.SelectCrossingWindow(ext.MinPoint, ext.MaxPoint, selFilter)
                If pmtSelRes.Status = PromptStatus.OK Then ' we found something.
                    Dim points As New Point3dCollection
                    Dim tmppnts As New Point3dCollection
                    mtxt.IntersectWith(tmpent, Intersect.OnBothOperands, tmppnts, IntPtr.Zero, IntPtr.Zero)
                    If tmppnts.Count > 0 Then 'we know we have an intersection this time.
                        tmppnts = Nothing
                        'if we new "points" here it shouldn't persist through our recursion?
                        points = New Point3dCollection
                    End If
                    mtxt.IntersectWith(tmpent, Intersect.OnBothOperands, points, IntPtr.Zero, IntPtr.Zero)
                    If points.Count > 0 Then ' we have an overlap?
                        Dim tmppnt As Point3d = mtxt.Location
                        pnt = points.Item(0)
                        mtxt.Location = GetBestDirectionForMove(tmppnt, pnt) 'New Point3d(tmppnt.X, tmppnt.Y + 1.0, tmppnt.Z)
                        mtxt.Location = DetermineOverlap(mtxt.ObjectId, tmpentId)
                        pnt = mtxt.Location
                    Else 'no overlap?
                        pnt = mtxt.Location
                        'Return pnt
                    End If
                    Return pnt
                Else
                    pnt = mtxt.Location
                    Return pnt
                End If
            Catch ex As Exception
                Active.WriteMessage(vbCrLf & "The error was: " & ex.Message)
            Finally
                tr.Commit()
                tr.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Gets an entity from an objectId
        ''' </summary>
        ''' <param name="objectId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetEntity(objectId As ObjectId) As Entity
            Dim ent As Entity
            Using tr As Transaction = Active.Database.TransactionManager.StartTransaction()
                ent = DirectCast(tr.GetObject(objectId, OpenMode.ForRead), Entity)
                tr.Commit()
            End Using
            Return ent
        End Function

        ''' <summary>
        ''' Gets the intersection point(s) of two entities and populates the intersectionPoints collection with the results.
        ''' </summary>
        ''' <param name="ent1"></param>
        ''' <param name="ent2"></param>
        ''' <param name="intersectionPoints"></param>
        ''' <param name="clash"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetIntersectionPointsOf(ent1 As Entity, ent2 As Entity, ByRef intersectionPoints As Point3dCollection, ByRef clash As Boolean) As Short
            
            Using tr As Transaction = Active.Database.TransactionManager.StartTransaction
                Dim oc1 = ent1.ObjectId.ObjectClass
                Dim oc2 = ent2.ObjectId.ObjectClass
                If (oc1.IsDerivedFrom(RXClass.GetClass(GetType(BlockReference)))) Then 'first entity is a blockreference
                    GetIntersectionPointsFromBlockEntities(ent1, ent2, intersectionPoints, clash)
                ElseIf (oc2.IsDerivedFrom(RXClass.GetClass(GetType(BlockReference)))) Then 'second entity is a blockreference
                    GetIntersectionPointsFromBlockEntities(ent2, ent1, intersectionPoints, clash)
                Else 'neither entities are blockreferences?
                    Dim tmppnts As New Point3dCollection
                    ent1.IntersectWith(ent2, Intersect.OnBothOperands, tmppnts, IntPtr.Zero, IntPtr.Zero)
                    If tmppnts.Count = 0 Then
                        clash = False
                    Else 'reset the clash value
                        clash = True
                    End If
                    ent1.IntersectWith(ent2, Intersect.OnBothOperands, intersectionPoints, IntPtr.Zero, IntPtr.Zero)
                End If
                tr.Commit()
            End Using
            Return (RTNORM)
        End Function

        ''' <summary>
        ''' determines the best way to move an object based on the insert point/origin,
        ''' and the angle to the first point of intersection.
        ''' </summary>
        ''' <param name="pnt">the point of intersection between the entities in question</param>
        ''' <param name="ent">the entity being tested</param>
        ''' <param name="blockent">the blockent for passing recursively back to GetIntersectionPointsFromBlockEntities</param>
        ''' <param name="intersectionPoints">the Point3dCollection we need to check each time.</param>
        ''' <remarks></remarks>
        Private Sub moveIt(pnt As Point3d, ent As Entity, blockent As Entity, ByRef intersectionPoints As Point3dCollection, ByRef clash As Boolean)
            'need to try and work out where the overlap occurs in relation to the origin/insert point of the entity.
            'Do Until intersectionPoints.Count = 0
            Dim tr As Transaction = Active.Database.TransactionManager.StartTransaction()
            Try
                Dim txt As DBText = TryCast(ent, DBText)
                If txt IsNot Nothing Then
                    'move it
                    txt.UpgradeOpen()
                    txt.Position = GetBestDirectionForMove(txt.Position, pnt)
                    'then check again.
                    GetIntersectionPointsOf(ent, blockent, intersectionPoints, clash)
                Else 'failed converting to DBText
                    Dim mtxt As MText = TryCast(ent, MText)
                    If mtxt IsNot Nothing Then
                        'move it
                        mtxt.UpgradeOpen()
                        mtxt.Location = GetBestDirectionForMove(mtxt.Location, pnt)
                        'then check again.
                        GetIntersectionPointsOf(ent, blockent, intersectionPoints, clash)
                    End If
                End If
            Catch ex As Exception

            Finally

                tr.Commit()
                tr.Dispose()
            End Try
            'Loop
        End Sub

        ''' <summary>
        ''' Calculates the best direction to move the mtext object based on the angle between the insertion point and intersectswith result.
        ''' </summary>
        ''' <param name="txtloc"></param>
        ''' <param name="pnt"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetBestDirectionForMove(txtloc As Point3d, pnt As Point3d) As Point3d
            Dim pt1 As Point2d = New Point2d(txtloc.X, txtloc.Y)
            Dim pt2 As Point2d = New Point2d(pnt.X, pnt.Y)
            Dim ang As Double = pt1.GetVectorTo(pt2).Angle ' ang is in radians so we need to modify our case values
            Select Case ang
                Case 0 To 1.57079633
                    txtloc = New Point3d(txtloc.X - 5, txtloc.Y + 0, txtloc.Z)
                Case 1.57079633 To 3.14159265
                    txtloc = New Point3d(txtloc.X, txtloc.Y - 5, txtloc.Z)
                Case 3.14159265 To 4.71238898
                    txtloc = New Point3d(txtloc.X, txtloc.Y - 5, txtloc.Z)
                Case 4.71238898 To 6.28318531
                    txtloc = New Point3d(txtloc.X - 5, txtloc.Y + 0, txtloc.Z)
            End Select
            Return txtloc
        End Function

        ''' <summary>
        ''' Gets the intersection points from Nested Block Entities.
        ''' </summary>
        ''' <param name="ent1"></param>
        ''' <param name="ent2"></param>
        ''' <param name="intersectionPoints"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetIntersectionPointsFromBlockEntities(ent1 As Entity, ent2 As Entity, ByRef intersectionPoints As Point3dCollection, ByRef clash As Boolean) As Short
            Using tr As Transaction = Active.Database.TransactionManager.StartTransaction
                Dim oc = ent2.ObjectId.ObjectClass
                If (oc.IsDerivedFrom(RXClass.GetClass(GetType(BlockReference)))) Then
                    Active.Editor.WriteMessage(vbCrLf & "Error - this routine cannot handle two blocks, sorry!")
                    Return (RTERROR)
                End If
                Dim pBlockRef = TryCast(ent1, BlockReference)
                If Not pBlockRef = Nothing Then
                    Dim entitySet As New DBObjectCollection
                    ent1.Explode(entitySet)
                    For i = 0 To entitySet.Count - 1
                        Dim tmpEntity As Entity = DirectCast(entitySet.Item(i), Entity)
                        If Not (tmpEntity.GetRXClass().DxfName = "ATTDEF") Then 'add any other entities that might fail to this line.
                            If Not GetIntersectionPointsOf(tmpEntity, ent2, intersectionPoints, clash) = RTNORM Then
                                Active.WriteMessage(vbCrLf & "Error getting objects!")
                            End If
                        End If
                    Next
                End If
                tr.Commit()
            End Using
            Return (RTNORM)
        End Function
#Region "BlockReplace.xml methods"

        ''' <summary>
        ''' Allows the user to add unknown borders to the BlockReplace.xml file.
        ''' </summary>
        ''' <remarks></remarks>
        <CommandMethod("GetPointsFromUser")> _
        Public Sub GetPointsFromUser()
            Dim prmptSelOpts As New PromptSelectionOptions()
            prmptSelOpts.MessageForAdding = "Select drawing frame block reference to capture Dumb Text coordinates from:"
            Dim prmptSelRes As PromptSelectionResult
            Dim rbfResult As ResultBuffer

            prmptSelRes = Active.Editor.GetSelection(prmptSelOpts)
            If prmptSelRes.Status <> PromptStatus.OK Then
                'the user didn't select anything.
                Exit Sub
            End If
            Dim selSet As SelectionSet
            selSet = prmptSelRes.Value
            rbfResult = New ResultBuffer( _
                        New TypedValue(LispDataType.SelectionSet, selSet))
            GetPointsFromUser(rbfResult)
        End Sub

        ''' <summary>
        ''' The worked for the GetPointsFromUser Method above.
        ''' Also callable from LISP.
        ''' </summary>
        ''' <param name="args">should contain only one block as a selectionset</param>
        ''' <remarks></remarks>
        <LispFunction("GetPointsFromUser")> _
        Public Sub GetPointsFromUser(args As ResultBuffer)
            If args Is Nothing Then
                Throw New ArgumentException("Requires one argument")
            End If
            Dim values As TypedValue() = args.AsArray
            If values.Length <> 1 Then
                Throw New ArgumentException("Wrong number of arguments")
            End If
            If values(0).TypeCode <> CInt(LispDataType.SelectionSet) Then
                Throw New ArgumentException("Bad argument type - requires a selection set")
            End If
            Dim ss As SelectionSet = DirectCast(values(0).Value, SelectionSet)
            Dim asmpath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim tmpblkname As String = ""
            Dim selEntID As ObjectId
            Dim selEntIDs As ObjectId()
            Using tmptrans As Transaction = Active.Database.TransactionManager.StartTransaction
                If ss.Count > 1 Then
                    For Each id As ObjectId In ss.GetObjectIds()
                        Dim tmpblkref As BlockReference = id.GetObject(OpenMode.ForRead)
                        If Not tmpblkref.Position = Point3d.Origin Then
                            Active.WriteMessage("This tool only works correctly if the block was inserted @ 0,0,0 coordinates!" & _
                                              vbCrLf & "Move the block to 0,0,0 and try again!")
                            Exit Sub
                        End If
                        If tmpblkref.Name.StartsWith("*") Then 'dynamic block
                            Dim tmpbtr As BlockTableRecord = tmpblkref.DynamicBlockTableRecord.GetObject(OpenMode.ForRead)
                            tmpblkname = tmpbtr.Name
                            selEntID = id
                        Else 'normal block
                            tmpblkname = tmpblkref.Name
                            selEntID = id
                        End If
                    Next
                Else
                    selEntIDs = ss.GetObjectIds() 'currently this doesn't play nice if there happens to be a second attributed block in the drawing!
                    selEntID = selEntIDs(0)
                    Dim tmpblkref As BlockReference = selEntID.GetObject(OpenMode.ForRead)
                    tmpblkname = tmpblkref.Name
                    If Not tmpblkref.Position = Point3d.Origin Then
                        Active.WriteMessage("This tool only works correctly if the block was inserted @ 0,0,0 coordinates!" & _
                                          vbCrLf & "Move the block to 0,0,0 and try again!")
                        Exit Sub
                    End If
                End If
            End Using

            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            DoneAddingPoints = False
            Do Until DoneAddingPoints
                Dim pko As New PromptKeywordOptions(vbLf & "Export Frame Points and Dumb Text Bounding Boxes?")

                pko.AllowNone = False
                pko.Keywords.Add("Yes")
                pko.Keywords.Add("No")
                pko.Keywords.[Default] = "Yes"

                Dim pkr As PromptResult = Active.Editor.GetKeywords(pko)
                pko = Nothing
                If pkr.Status <> PromptStatus.OK OrElse pkr.StringResult = "No" Then
                    Return
                End If

                Dim pPtRes As PromptPointResult
                If pkr.Status = PromptStatus.OK Then
                    Using acTrans As Transaction = Active.Database.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable
                        '' Open Model space for write
                        acBlkTbl = acTrans.GetObject(Active.Database.BlockTableId, OpenMode.ForRead)
                        Dim serializer As New XmlSerializer(GetType(frames))
                        Dim f As New frames
                        Dim frame As New framesDrawingFrame
                        '' Prompt for the start point
                        pPtOpts.Message = vbLf & "Enter the bottom left corner point of the SearchBox: "
                        pPtRes = Active.Editor.GetPoint(pPtOpts)
                        Dim ptStart As Point3d = pPtRes.Value

                        '' Exit if the user presses ESC or cancels the command
                        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

                        '' Prompt for the end point
                        pPtOpts.Message = vbLf & "Enter the top right corner point of the SearchBox: "
                        pPtOpts.UseBasePoint = True
                        pPtOpts.BasePoint = ptStart
                        pPtRes = Active.Editor.GetPoint(pPtOpts)
                        Dim ptEnd As Point3d = pPtRes.Value

                        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub
                        Dim pnt3d1 As New framesDrawingFrameSearchBoxBoundsPoint3d With {.X = ptStart.X, .Y = ptStart.Y, .Z = ptStart.Z, .name = "p1"}
                        Dim pnt3d2 As New framesDrawingFrameSearchBoxBoundsPoint3d With {.X = ptStart.X, .Y = ptEnd.Y, .Z = ptStart.Z, .name = "p2"}
                        Dim pnt3d3 As New framesDrawingFrameSearchBoxBoundsPoint3d With {.X = ptEnd.X, .Y = ptEnd.Y, .Z = ptStart.Z, .name = "p3"}
                        Dim pnt3d4 As New framesDrawingFrameSearchBoxBoundsPoint3d With {.X = ptEnd.X, .Y = ptStart.Y, .Z = ptStart.Z, .name = "p4"}
                        frame.SearchBoxBounds.Points.Add(pnt3d1)
                        frame.SearchBoxBounds.Points.Add(pnt3d2)
                        frame.SearchBoxBounds.Points.Add(pnt3d3)
                        frame.SearchBoxBounds.Points.Add(pnt3d4)
                        'End If
                        KeepAddingAttRefs = True
                        Do Until KeepAddingAttRefs = False
                            pko = New PromptKeywordOptions("Add Dumb Text Bounding Boxes?")
                            'pko.Message = "Add Dumb Text Bounding Boxes?"
                            pko.Keywords.Add("Yes")
                            pko.Keywords.Add("No")
                            pko.Keywords.[Default] = "Yes"
                            pkr = Active.Editor.GetKeywords(pko)
                            pko = Nothing
                            If pkr.Status = PromptStatus.OK And Not pkr.StringResult = "No" Then
                                Dim attref As New framesDrawingFrameSearchBoxBoundsAttref
                                '' Prompt for the start point
                                Dim tmppPtOpts As PromptPointOptions = New PromptPointOptions(vbLf & "Enter the bottom left point of the line: ")
                                Dim tmppPtRes As PromptPointResult
                                tmppPtRes = Active.Editor.GetPoint(tmppPtOpts)
                                Dim tmpptStart As Point3d = tmppPtRes.Value

                                '' Exit if the user presses ESC or cancels the command
                                If tmppPtRes.Status = PromptStatus.Cancel Then Exit Do

                                '' Prompt for the end point
                                tmppPtOpts.Message = vbLf & "Enter the top right point of the line: "
                                tmppPtOpts.UseBasePoint = True
                                tmppPtOpts.BasePoint = tmpptStart
                                tmppPtRes = Active.Editor.GetPoint(tmppPtOpts)
                                tmppPtOpts = Nothing
                                Dim tmpptEnd As Point3d = tmppPtRes.Value
                                If tmppPtRes.Status = PromptStatus.Cancel Then Exit Do
                                Dim pStrAttRefOpts As PromptKeywordOptions = New PromptKeywordOptions(vbCrLf & "Enter Attribute Name: ")
                                pStrAttRefOpts.AllowArbitraryInput = True
                                pStrAttRefOpts.Keywords.Add("SUBCONTRACTOR")
                                pStrAttRefOpts.Keywords.Add("CADREF")
                                pStrAttRefOpts.Keywords.Add("RESPONSIBLEAUTHORITY")
                                pStrAttRefOpts.Keywords.Add("TOLERANCES")
                                pStrAttRefOpts.Keywords.Add("SURFACETEXTURE")
                                pStrAttRefOpts.Keywords.Add("PROTECTIVEFINISH")
                                pStrAttRefOpts.Keywords.Add("MATERIAL")
                                pStrAttRefOpts.Keywords.Add("USEDON")
                                pStrAttRefOpts.Keywords.Add("REVISIONS")
                                pStrAttRefOpts.Keywords.Add("SCALE")
                                pStrAttRefOpts.Keywords.Add("DIMSIN")



                                Dim pStrAttRefRes As PromptResult = Active.Editor.GetKeywords(pStrAttRefOpts)

                                Select Case pStrAttRefRes.StringResult
                                    Case "CADREF"
                                        attref.name = "CAD_REF"
                                        attref.AttributeName = "CAD REF"
                                    Case "RESPONSIBLEAUTHORITY"
                                        attref.name = "RESPONSIBLE_AUTHORITY"
                                        attref.AttributeName = "RESPONSIBLE AUTHORITY"
                                    Case "SURFACETEXTURE"
                                        attref.name = "SURFACE_TEXTURE"
                                        attref.AttributeName = "SURFACE TEXTURE"
                                    Case "PROTECTIVEFINISH"
                                        attref.name = "PROTECTIVE_FINISH"
                                        attref.AttributeName = "PROTECTIVE FINISH"
                                    Case "TOLERANCES" 'tolerances & material
                                        attref.name = pStrAttRefRes.StringResult
                                        attref.AttributeName = "TOLERANCE"
                                    Case "USEDON"
                                        attref.name = "USED_ON"
                                    Case "REVISIONS"
                                        attref.name = "REVISIONS"
                                    Case "SCALE"
                                        attref.name = "SCALE"
                                        attref.AttributeName = "SCALE"
                                    Case "DIMSIN"
                                        attref.name = "DIMSIN"
                                        attref.AttributeName = "DIMS IN"
                                    Case "SUBCONTRACTOR"
                                        attref.name = "SUBCONTRACTOR"
                                        If tmpblkname = "SB-IL_996-5.2(block)" Then attref.AttributeName = "SUB CONTRACTOR"
                                        attref.AttributeName = "SUB-CONTRACTOR"
                                    Case Else
                                        attref.name = pStrAttRefRes.StringResult
                                        attref.AttributeName = pStrAttRefRes.StringResult
                                End Select
                                pStrAttRefOpts = Nothing

                                Dim attrefpnt1 As New framesDrawingFrameSearchBoxBoundsAttrefPoint3d With {.X = tmpptStart.X, .Y = tmpptStart.Y, .Z = tmpptStart.Z, .name = "p1"}
                                Dim attrefpnt2 As New framesDrawingFrameSearchBoxBoundsAttrefPoint3d With {.X = tmpptStart.X, .Y = tmpptEnd.Y, .Z = tmpptStart.Z, .name = "p2"}
                                Dim attrefpnt3 As New framesDrawingFrameSearchBoxBoundsAttrefPoint3d With {.X = tmpptEnd.X, .Y = tmpptEnd.Y, .Z = tmpptEnd.Z, .name = "p3"}
                                Dim attrefpnt4 As New framesDrawingFrameSearchBoxBoundsAttrefPoint3d With {.X = tmpptEnd.X, .Y = tmpptStart.Y, .Z = tmpptStart.Z, .name = "p4"}
                                attref.point3d.Add(attrefpnt1)
                                attref.point3d.Add(attrefpnt2)
                                attref.point3d.Add(attrefpnt3)
                                attref.point3d.Add(attrefpnt4)
                                frame.SearchBoxBounds.Attrefs.Add(attref)
                                attref = Nothing
                            ElseIf pkr.Status <> PromptStatus.OK OrElse pkr.StringResult = "No" Then
                                KeepAddingAttRefs = False
                            End If
                        Loop
                        If tmpblkname Like "*5.2(block)" Then
                            KeepAddingCrossingLines = False
                            Do Until KeepAddingCrossingLines
                                pko = New PromptKeywordOptions("Add Crossing Box Lines?")
                                pko.Keywords.Add("Yes")
                                pko.Keywords.Add("No")
                                pko.Keywords.[Default] = "Yes"
                                pkr = Active.Editor.GetKeywords(pko)
                                pko = Nothing
                                If pkr.Status = PromptStatus.OK And Not pkr.StringResult = "No" Then
                                    Dim tmppPtOpts As PromptPointOptions = New PromptPointOptions(vbLf & "Enter the top left point of the line: ")
                                    Dim tmppPtRes As PromptPointResult
                                    'Dim lines As New framesDrawingFrameLine()
                                    Dim line As New framesDrawingFrameLine
                                    tmppPtRes = Active.Editor.GetPoint(tmppPtOpts)
                                    Dim tmpptStart As Point3d = tmppPtRes.Value

                                    '' Exit if the user presses ESC or cancels the command
                                    If tmppPtRes.Status = PromptStatus.Cancel Then Exit Do

                                    '' Prompt for the end point
                                    tmppPtOpts.Message = vbLf & "Enter the bottom right point of the line: "
                                    tmppPtOpts.UseBasePoint = True
                                    tmppPtOpts.BasePoint = tmpptStart
                                    tmppPtRes = Active.Editor.GetPoint(tmppPtOpts)
                                    tmppPtOpts = Nothing
                                    Dim tmpptEnd As Point3d = tmppPtRes.Value
                                    If tmppPtRes.Status = PromptStatus.Cancel Then Exit Do
                                    Dim pStrAttRefOpts As PromptKeywordOptions = New PromptKeywordOptions(vbCrLf & "Enter Line Name: ")
                                    pStrAttRefOpts.AllowArbitraryInput = True
                                    pStrAttRefOpts.Keywords.Add("DRAWINGOFFICEREF")
                                    pStrAttRefOpts.Keywords.Add("CADREF")
                                    pStrAttRefOpts.Keywords.Add("RESPONSIBLEAUTHORITY")
                                    pStrAttRefOpts.Keywords.Add("TOLERANCE")
                                    pStrAttRefOpts.Keywords.Add("SURFACETEXTURE")
                                    pStrAttRefOpts.Keywords.Add("PROTECTIVEFINISH")
                                    pStrAttRefOpts.Keywords.Add("MATERIAL")
                                    pStrAttRefOpts.Keywords.Add("SUBCONTRACTOR")
                                    pStrAttRefOpts.Keywords.Add("SCALE")
                                    pStrAttRefOpts.Keywords.Add("DIMSIN")

                                    Dim pStrAttRefRes As PromptResult = Active.Editor.GetKeywords(pStrAttRefOpts)

                                    Select Case pStrAttRefRes.StringResult
                                        Case "DRAWINGOFFICEREF"
                                            line.name = "DRAWING OFFICE REF"
                                        Case "CADREF"
                                            line.name = "CAD REF"
                                        Case "RESPONSIBLEAUTHORITY"
                                            line.name = "RESPONSIBLE AUTHORITY"
                                        Case "SURFACETEXTURE"
                                            line.name = "SURFACE TEXTURE"
                                        Case "SUBCONTRACTOR"
                                            If tmpblkname = "SB-IL_996-5.2(block)" Then line.name = "SUB CONTRACTOR"
                                            line.name = "SUB-CONTRACTOR"
                                        Case "PROTECTIVEFINISH"
                                            line.name = "PROTECTIVE FINISH"
                                        Case "DIMSIN"
                                            line.name = "DIMS IN"
                                        Case Else 'material, tolerances, scale
                                            line.name = pStrAttRefRes.StringResult
                                    End Select
                                    pStrAttRefOpts = Nothing
                                    line.Layer = "FRAME"
                                    line.LineStart = New framesDrawingFrameLineLineStart With {.X = tmpptStart.X, .Y = tmpptStart.Y, .Z = tmpptStart.Z}
                                    line.LineEnd = New framesDrawingFrameLineLineEnd With {.X = tmpptEnd.X, .Y = tmpptEnd.Y, .Z = tmpptEnd.Z}
                                    frame.Lines.Add(line)
                                    line = Nothing
                                ElseIf pkr.Status <> PromptStatus.OK OrElse pkr.StringResult = "No" Then
                                    KeepAddingCrossingLines = True
                                End If
                            Loop
                        End If

                        'Application.ShowAlertDialog("The name entered was: " & pStrRes.StringResult)
                        frame.name = tmpblkname
                        f.DrawingFrame.Add(frame)
                        'f.SaveToFile(asmpath & "\Resources\test.xml")
                        Active.WriteMessage(f.ToString)
                        Dim fs As FileStream
                        If Not File.Exists(asmpath & "\Resources\" & frame.name & ".xml") Then
                            fs = New FileStream(asmpath & "\Resources\" & frame.name & ".xml", FileMode.Create)
                        Else
                            fs = New FileStream(asmpath & "\Resources\" & frame.name & ".xml", FileMode.Append)
                        End If
                        Dim writer As New XmlTextWriter(fs, Encoding.Unicode)
                        writer.Formatting = Formatting.Indented
                        serializer.Serialize(writer, f)
                        writer.Close()
                        '' Commit the changes and dispose of the transaction even though we aren't adding anything
                        acTrans.Commit()
                    End Using
                ElseIf pkr.Status <> PromptStatus.OK OrElse pkr.StringResult = "No" Then
                    DoneAddingPoints = True
                ElseIf pkr.Status = PromptStatus.Cancel Then
                    Exit Sub
                End If
            Loop
        End Sub

        ''' <summary>
        ''' Deprecated in favour of the above because this one relies on the EverythingisAWEsome borders having fully enclosed text boxes
        ''' which of course they don't!
        ''' Copied from here http://docs.autodesk.com/ACD/2010/ENU/AutoCAD%20.NET%20Developer's%20Guide/index.html?url=WS1a9193826455f5ff2566ffd511ff6f8c7ca-4217.htm,topicNumber=d0e10700
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub GetPointsFromUser(ByVal pnts As Point3dCollection)
            Dim pPtRes As PromptPointResult
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

            '' Prompt for the start point
            pPtOpts.Message = vbLf & "Enter the start point of the line: "
            pPtRes = Active.Editor.GetPoint(pPtOpts)
            Dim ptStart As Point3d = pPtRes.Value

            '' Exit if the user presses ESC or cancels the command
            If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

            '' Prompt for the end point
            pPtOpts.Message = vbLf & "Enter the end point of the line: "
            pPtOpts.UseBasePoint = True
            pPtOpts.BasePoint = ptStart
            pPtRes = Active.Editor.GetPoint(pPtOpts)
            Dim ptEnd As Point3d = pPtRes.Value

            If pPtRes.Status = PromptStatus.Cancel Then Exit Sub

            '' Start a transaction
            Using acTrans As Transaction = Active.Database.TransactionManager.StartTransaction()

                Dim acBlkTbl As BlockTable
                '' Open Model space for write
                acBlkTbl = acTrans.GetObject(Active.Database.BlockTableId, OpenMode.ForRead)
                Dim serializer As New XmlSerializer(GetType(frames))
                Dim f As New frames
                Dim frame As New framesDrawingFrame
                Dim pnt3d1 As New framesDrawingFrameSearchBoxBoundsPoint3d
                With pnt3d1
                    .X = ptStart.X
                    .Y = ptStart.Y
                    .Z = ptStart.Z
                    .name = "p1"
                End With
                Dim pnt3d2 As New framesDrawingFrameSearchBoxBoundsPoint3d
                With pnt3d2
                    .X = ptStart.X
                    .Y = ptEnd.Y
                    .Z = ptStart.Z
                    .name = "p2"
                End With
                Dim pnt3d3 As New framesDrawingFrameSearchBoxBoundsPoint3d
                With pnt3d3
                    .X = ptEnd.X
                    .Y = ptEnd.Y
                    .Z = ptStart.Z
                    .name = "p3"
                End With
                Dim pnt3d4 As New framesDrawingFrameSearchBoxBoundsPoint3d
                With pnt3d4
                    .X = ptEnd.X
                    .Y = ptStart.Y
                    .Z = ptStart.Z
                    .name = "p4"
                End With
                frame.SearchBoxBounds.Points.Add(pnt3d1)
                frame.SearchBoxBounds.Points.Add(pnt3d2)
                frame.SearchBoxBounds.Points.Add(pnt3d3)
                frame.SearchBoxBounds.Points.Add(pnt3d4)
                Dim pStrOpts As PromptStringOptions = New PromptStringOptions(vbLf & "Enter the frame name: ")
                pStrOpts.AllowSpaces = True
                Dim pStrRes As PromptResult = Active.Editor.GetString(pStrOpts)
                'Application.ShowAlertDialog("The name entered was: " & pStrRes.StringResult)
                frame.name = pStrRes.StringResult
                f.DrawingFrame.Add(frame)
                Dim asmpath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                'f.SaveToFile(asmpath & "\Resources\test.xml")
                Active.WriteMessage(f.ToString)
                Dim fs As FileStream
                If Not File.Exists(asmpath & "\Resources\" & pStrRes.StringResult & ".xml") Then
                    fs = New FileStream(asmpath & "\Resources\" & pStrRes.StringResult & ".xml", FileMode.Create)
                Else
                    fs = New FileStream(asmpath & "\Resources\" & pStrRes.StringResult & ".xml", FileMode.Append)
                End If
                Dim writer As New XmlTextWriter(fs, Encoding.Unicode)
                serializer.Serialize(writer, f)
                writer.Close()

                '' Commit the changes and dispose of the transaction even though we aren't adding anything
                acTrans.Commit()
            End Using
        End Sub
#End Region
        ''' <summary>
        ''' An example class showing how to filter for lines in a selectionset based upon a known rectangle "search area"
        ''' Copied in from http://adndevblog.typepad.com/autocad/2012/05/use-selectcrossingpolygon-to-select-entities-in-view-other-than-plan.html
        ''' </summary>
        ''' <remarks></remarks>
        <CommandMethod("SEL")> _
        Public Sub MySelection()
            Dim p1 As New Point3d(10.0, 10.0, 0.0)
            Dim p2 As New Point3d(10.0, 11.0, 0.0)
            Dim p3 As New Point3d(11.0, 11.0, 0.0)
            Dim p4 As New Point3d(11.0, 10.0, 0.0)

            Dim pntCol As New Point3dCollection()
            pntCol.Add(p1)
            pntCol.Add(p2)
            pntCol.Add(p3)
            pntCol.Add(p4)

            Dim pmtSelRes As PromptSelectionResult = Nothing

            Dim typedVal As TypedValue() = New TypedValue(0) {}
            typedVal(0) = New TypedValue(CInt(DxfCode.Start), "Line")

            Dim selFilter As New SelectionFilter(typedVal)
            pmtSelRes = Active.Editor.SelectCrossingPolygon(pntCol, selFilter)

            If pmtSelRes.Status = PromptStatus.OK Then
                Active.WriteMessage("Entities found " & pmtSelRes.Value.GetObjectIds().Count.ToString())
            Else
                Active.WriteMessage(vbLf & "Did not find entities")
            End If
        End Sub

        ''' <summary>
        ''' An example class showing how to prompt for user selection and changes the coloUr of anything selected.
        ''' Copied in from http://docs.autodesk.com/ACD/2013/ENU/index.html?url=files/GUID-CBECEDCF-3B4E-4DF3-99A0-47103D10DADD.htm,topicNumber=d30e724932
        ''' </summary>
        ''' <remarks></remarks>
        <CommandMethod("SelectObjectsOnscreen")> _
        Public Sub SelectObjectsOnscreen()
            '' Start a transaction
            Using acTrans As Transaction = Active.Database.TransactionManager.StartTransaction()

                '' Request for objects to be selected in the drawing area
                Dim acSSPrompt As PromptSelectionResult = Active.Editor.GetSelection()

                '' If the prompt status is OK, objects were selected
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value

                    '' Step through the objects in the selection set
                    For Each acSSObj As SelectedObject In acSSet
                        '' Check to make sure a valid SelectedObject object was returned
                        If Not IsDBNull(acSSObj) Then
                            '' Open the selected object for write
                            Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, _
                                                                    OpenMode.ForWrite)

                            If Not IsDBNull(acEnt) Then
                                '' Change the object's color to Green
                                acEnt.ColorIndex = 3
                            End If
                        End If
                    Next

                    '' Save the new object to the database
                    acTrans.Commit()
                End If

                '' Dispose of the transaction
            End Using
        End Sub

        ''' <summary>
        ''' Inserts a block from a wblocked .dwg file.
        ''' </summary>
        ''' <param name="DatabaseIn">The source database</param>
        ''' <param name="BTRToAddTo">the BlockTableRecord we're adding this BlockReference to.</param>
        ''' <param name="InsPt">the Point3d insertion point for the new block</param>
        ''' <param name="BlockName">the name of the block to add</param>
        ''' <param name="XScale">the Xscale of the existing block</param>
        ''' <param name="YScale">the Yscale of the existing block</param>
        ''' <param name="ZScale">the Zscale of the existing block</param>
        ''' <returns>returns the objectId of the newly inserted blockreference</returns>
        ''' <remarks></remarks>
        Public Function InsertBlock(ByVal DatabaseIn As Database, _
                        ByVal btrtoaddto As String, _
                        ByVal inspt As Geometry.Point3d, _
                        ByVal blockname As String, _
                        ByVal xscale As Double, _
                        ByVal yscale As Double, _
                        ByVal zscale As Double) As DatabaseServices.ObjectId
            Using myTrans As Transaction = DatabaseIn.TransactionManager.StartTransaction
                Dim myBlockTable As BlockTable = DatabaseIn.BlockTableId.GetObject(OpenMode.ForRead)
                'If the suppplied Block Name is not
                'in the specified Database, get out gracefully.
                If myBlockTable.Has(blockname) = False Then
                    Return Nothing
                End If
                'If the specified BlockTableRecord does not exist,
                'get out gracefully
                If myBlockTable.Has(btrtoaddto) = False Then
                    Return Nothing
                End If
                Dim myBlockDef As BlockTableRecord = myBlockTable(blockname).GetObject(OpenMode.ForRead)
                Dim myBlockTableRecord As BlockTableRecord = myBlockTable(btrtoaddto).GetObject(OpenMode.ForWrite)
                'Create a new BlockReference
                Dim myBlockRef As New BlockReference(inspt, myBlockDef.Id)
                'Set the scale factors
                myBlockRef.ScaleFactors = New Geometry.Scale3d(xscale, yscale, zscale)
                'Add the new BlockReference to the specified BlockTableRecord
                myBlockTableRecord.AppendEntity(myBlockRef)
                'Add the BlockReference to the BlockTableRecord.
                myTrans.AddNewlyCreatedDBObject(myBlockRef, True)
                Dim myAttColl As DatabaseServices.AttributeCollection = myBlockRef.AttributeCollection
                'Find Attributes and add them to the AttributeCollection
                'of the BlockReference
                For Each myEntID As ObjectId In myBlockDef
                    Dim myEnt As Entity = myEntID.GetObject(OpenMode.ForRead)
                    If TypeOf myEnt Is DatabaseServices.AttributeDefinition Then
                        Dim myAttDef As DatabaseServices.AttributeDefinition = myEnt
                        Dim myAttRef As New DatabaseServices.AttributeReference
                        myAttRef.SetAttributeFromBlock(myAttDef, myBlockRef.BlockTransform)
                        myAttColl.AppendAttribute(myAttRef)
                        myTrans.AddNewlyCreatedDBObject(myAttRef, True)
                    End If
                Next
                myTrans.Commit()
                Return myBlockRef.Id
            End Using
        End Function

        ''' <summary>
        ''' Replaces an existing BlockReference with the latest version
        ''' copied from: http://adndevblog.typepad.com/autocad/2012/05/redefining-a-block.html
        ''' </summary>
        ''' <param name="objectId">the objectId of the block to be replaced.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ReplaceBlock(objectId As ObjectId) As ObjectId
            Dim Tx As Transaction = Active.Database.TransactionManager.StartTransaction()
            Dim blkDb As Database = New Database(False, True)
            Try
                Dim tmpblkref As BlockReference = Tx.GetObject(objectId, OpenMode.ForRead)
                blkDb.ReadDwgFile("C:\LEGACY VAULT WORKING FOLDER\Designs\AWE\FSTF\AWE DRAWING FRAMES\" & tmpblkref.Name & ".dwg", System.IO.FileShare.Read, True, "")
                Dim blockTable As BlockTable = Tx.GetObject(Active.Database.BlockTableId, OpenMode.ForRead, False, True)
                Dim btrId As ObjectId = Active.Database.Insert(tmpblkref.Name, blkDb, True)
                If btrId <> objectId.Null Then
                    Dim btr As BlockTableRecord = Tx.GetObject(btrId, OpenMode.ForRead, False, True)
                    Dim brefIds As ObjectIdCollection = btr.GetBlockReferenceIds(False, True)
                    For Each id As ObjectId In brefIds
                        Dim bref As BlockReference =
                        Tx.GetObject(id, OpenMode.ForWrite, False, True)
                        bref.RecordGraphicsModified(True)
                    Next
                End If
            Catch ex As Exception
                Tx.Commit()
                Tx.Dispose()
                blkDb.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' The "Manual" Command for SortRevisions
        ''' </summary>
        ''' <remarks></remarks>
        <CommandMethod("SortRevisions")> _
        Public Sub SortRevisions()
            If revisions Is Nothing Then
                revisions = New List(Of Revision)()
            End If
            Dim prmptSelOpts As New PromptSelectionOptions()
            prmptSelOpts.SingleOnly = True
            prmptSelOpts.SinglePickInSpace = True
            Dim prmptSelRes As PromptSelectionResult
            Dim prmptRNStrOpts As PromptStringOptions = New PromptStringOptions(vbLf + "Enter RN")
            prmptRNStrOpts.AllowSpaces = False
            prmptRNStrOpts.DefaultValue = "RN240129"
            prmptRNStrOpts.UseDefaultValue = True
            Dim prmptDateStrOpts As PromptStringOptions = New PromptStringOptions(vbLf + "Enter desired date")
            prmptDateStrOpts.DefaultValue = System.DateTime.Now.ToString("dd-MMM-yy")
            prmptDateStrOpts.UseDefaultValue = True
            Dim prdate As PromptResult = Active.Editor.GetString(prmptDateStrOpts)
            If Not prdate.Status = PromptStatus.OK Then
                Exit Sub
            End If
            Dim pr As PromptResult = Active.Editor.GetString(prmptRNStrOpts)
            If Not pr.Status = PromptStatus.OK Then
                Exit Sub
            End If
            prmptSelRes = Active.Editor.GetSelection(prmptSelOpts)
            If prmptSelRes.Status <> PromptStatus.OK Then
                'the user didn't select anything.
                Exit Sub
            End If
            Dim selSet As SelectionSet
            selSet = prmptSelRes.Value
            'Package the results...
            Dim rbfResult As ResultBuffer
            rbfResult = New ResultBuffer( _
                        New TypedValue(LispDataType.SelectionSet, selSet), _
                        New TypedValue(LispDataType.Text, pr.StringResult), _
                        New TypedValue(LispDataType.Text, prdate.StringResult))
            SortRevs(rbfResult)
        End Sub

        ''' <summary>
        ''' The Lisp-only version of SortRevisions.
        ''' </summary>
        ''' <param name="args">Expects at least one object in a selectionset</param>
        ''' <remarks>requires the arguments of RN###### & a release date in dd-MMM-yy format</remarks>
        <LispFunction("SortRevs")> _
        Public Sub SortRevs(args As ResultBuffer)
            Try
                If revisions Is Nothing Then
                    revisions = New List(Of Revision)()
                End If
                'Microsoft.VisualBasic.MsgBox("Attach to Process now!")
                If args Is Nothing Then
                    Throw New ArgumentException("Requires one argument")
                End If
                Dim values As TypedValue() = args.AsArray
                If values.Length <> 3 Then
                    Throw New ArgumentException("Wrong number of arguments")
                End If
                If values(0).TypeCode <> CInt(LispDataType.SelectionSet) Then
                    Throw New ArgumentException("Bad argument type - requires a selection set")
                End If
                If values(1).TypeCode <> CInt(LispDataType.Text) Then
                    Throw New ArgumentException("Bad argument type - requires a text string in the RN###### format")
                End If
                If values(2).TypeCode <> CInt(LispDataType.Text) Then
                    Throw New ArgumentException("Bad argument type - requires a date text string in the dd-MMM-yy format")
                End If
                Dim ss As SelectionSet = DirectCast(values(0).Value, SelectionSet)
                Dim ReleaseNote As String = DirectCast(values(1).Value, String)
                Dim ReleaseDate As String = DirectCast(values(2).Value, String)
                Dim strdrawingfilename As String = Application.GetSystemVariable("DWGNAME")
                Dim drawingpath As String = Application.GetSystemVariable("DWGPREFIX")
                Dim supercededPath As String
                If Not drawingpath.ToUpper.Contains("CURRENT") Then 'file is in the root of the folder and we need to create the correct folders for it.
                    IO.Directory.CreateDirectory(drawingpath & "CURRENT\")
                    IO.Directory.CreateDirectory(drawingpath & "SUPERCEDED\")
                    supercededPath = drawingpath & "SUPERCEDED\"
                    drawingpath = drawingpath & "CURRENT\"
                Else
                    supercededPath = GetParentDirectory(drawingpath, 2) & "\Superceded\"
                End If
                Dim originalfilename As String = Active.Database.Filename
                Dim dwgnum As String = ""
                Dim shtnum As String = ""
                Dim IssueChar As String = ""
                Dim selEntID As ObjectId = Nothing
                Dim selEntIDs As ObjectId()
                Dim tmpblkname As String
                Dim ext As Extents3d = If(CShort(Application.GetSystemVariable("cvport")) = 1, New Extents3d(Active.Database.Pextmin, Active.Database.Pextmax), New Extents3d(Active.Database.Extmin, Active.Database.Extmax))
                Using doclock As DocumentLock = Active.Document.LockDocument 'lock the document whilst we edit it.
                    Using tr As Transaction = Active.Database.TransactionManager.StartTransaction
                        If ss.Count > 1 Then
                            'get the objectId from the blockname
                            For Each id As ObjectId In ss.GetObjectIds()
                                Dim tmpblkref As BlockReference = id.GetObject(OpenMode.ForWrite)
                                If tmpblkref.Name.StartsWith("*") Then 'dynamic block
                                    Dim tmpbtr As BlockTableRecord = tmpblkref.DynamicBlockTableRecord.GetObject(OpenMode.ForRead)
                                    tmpblkname = tmpbtr.Name
                                    If tmpblkname Like "*5.2(block)" Then
                                        selEntID = id
                                        Exit For
                                    End If
                                Else 'normal block
                                    tmpblkname = tmpblkref.Name
                                    If tmpblkname Like "*5.2(block)" Then
                                        selEntID = id
                                        Exit For
                                    End If
                                End If
                            Next
                        Else
                            selEntIDs = ss.GetObjectIds()
                            selEntID = selEntIDs(0)
                        End If
                        If selEntID = Nothing Then 'we're not running this tool on a "*5.2(block)" named drawing frame.
                            Active.WriteMessage("You need to make sure the block has been replaced before running this tool" + vbCrLf + "Exiting...")
                            Exit Sub
                        End If
                        Dim blkref As BlockReference = selEntID.GetObject(OpenMode.ForWrite)
                        If Not blkref.Name Like ("*5.2(block)") Then
                            Active.WriteMessage("You need to make sure the block has been replaced before running this tool" + vbCrLf + "Exiting...")
                            Exit Sub
                        End If
                        Dim attcoll As AttributeCollection = blkref.AttributeCollection
                        'commented out this for now - can return to this later!
                        'Dim regex = New Regex("(?<chngnote>CHANGE NO \d*)\|(?<date>DATE \d*)\|(?<issue>ISSUE \d*)\|(?<dwgnum>DRAWING NUMBER 1)\|(?<shtnum>SHEET NUMBER 1)")
                        Dim AllSearchTerm As Regex = New Regex("(.* )(\d*)")
                        Dim ChangeNoteSearchTerm As Regex = New Regex("(CHANGE NO \d*)")
                        Dim DateSearchTerm As Regex = New Regex("(DATE \d*)")
                        Dim IssueSearchTerm As Regex = New Regex("(ISSUE \d*)")
                        Dim DrawingNumberSearchTerm As Regex = New Regex("(DRAWING NUMBER 1)")
                        Dim SheetNumberSearchTerm As Regex = New Regex("(SHEET NUMBER 1)")
                        Dim queryMatchingAll _
                            = From a As ObjectId In attcoll
                              Let b As AttributeReference = a.GetObject(OpenMode.ForRead)
                              Let n As String() = b.Tag.Split(New Char() {" "})
                              Let matches = AllSearchTerm.Matches(b.Tag)
                              Where (matches.Count > 0)
                              Order By n(0)
                              Group By groupKey = n(n.Length - 1)
                              Into groupName = Group

                        For Each gGroup In queryMatchingAll
                            'Active.WriteMessage(vbLf + "querymatchingall: " + (gGroup.groupKey) + vbLf)
                            Dim rev As New Revision
                            Dim chngNote As AttRef = Nothing
                            Dim issDate As AttRef = Nothing
                            Dim issue As AttRef = Nothing


                            'chngNote = GetAttributeReferenceFromQMA(ChangeNoteSearchTerm, gGroup)
                            'issDate = GetAttributeReferenceFromQMA(DateSearchTerm, gGroup)
                            'issue = GetAttributeReferenceFromQMA(IssueSearchTerm, gGroup)
                            'dwgnum = GetAttributeReferenceFromQMA(DrawingNumberSearchTerm, gGroup).attRefText.Replace("/", "-")
                            'shtnum = GetAttributeReferenceFromQMA(SheetNumberSearchTerm, gGroup).attRefText

                            Dim CN = (From item In gGroup.groupName
                                                Let matches = ChangeNoteSearchTerm.Matches(item.b.Tag)
                                                Where (matches.Count > 0)
                                                Select item.b).FirstOrDefault()
                            If Not CN = Nothing Then
                                chngNote = New AttRef
                                chngNote.attrefId = CN.ObjectId
                                chngNote.attRefTag = CN.Tag
                                chngNote.attRefText = CN.TextString
                            End If
                            Dim DT = (From item In gGroup.groupName
                                      Let matches = DateSearchTerm.Matches(item.b.Tag)
                                      Where (matches.Count > 0)
                                      Select item.b).FirstOrDefault()
                            If Not DT = Nothing Then
                                issDate = New AttRef
                                issDate.attrefId = DT.ObjectId
                                issDate.attRefTag = DT.Tag
                                issDate.attRefText = DT.TextString
                            End If
                            Dim RV = (From item In gGroup.groupName
                                      Let matches = IssueSearchTerm.Matches(item.b.Tag)
                                      Where (matches.Count > 0)
                                      Select item.b).FirstOrDefault()
                            If Not RV = Nothing Then
                                issue = New AttRef
                                issue.attrefId = RV.ObjectId
                                issue.attRefTag = RV.Tag
                                issue.attRefText = RV.TextString
                            End If
                            Dim dnum = (From item In gGroup.groupName
                                       Let matches = DrawingNumberSearchTerm.Matches(item.b.Tag)
                                       Where matches.Count > 0
                                       Select item.b).FirstOrDefault()
                            If Not dnum = Nothing Then
                                dwgnum = dnum.TextString.Replace("/", "-")
                            End If
                            Dim snum = (From item In gGroup.groupName
                                       Let matches = SheetNumberSearchTerm.Matches(item.b.Tag)
                                       Where matches.Count > 0
                                       Select item.b).FirstOrDefault()
                            If Not snum = Nothing Then
                                shtnum = snum.TextString
                            End If
                            If Not chngNote Is Nothing And Not issDate Is Nothing And Not issue Is Nothing Then
                                rev.RevChangeNote = chngNote
                                rev.RevDate = issDate
                                rev.RevIssue = issue
                                revisions.Add(rev)
                            End If
                            'Active.WriteMessage("Captured: " + revisions.Count.ToString() + " revision rows!")
                            chngNote = Nothing
                            issDate = Nothing
                            issue = Nothing
                        Next
                        queryMatchingAll = Nothing
                        'First check for non-standard revisions.
                        For i As Integer = 0 To revisions.Count - 1
                            If MatchesOddRevision(revisions.Item(i).RevIssue.attRefText) Then 'blank off any with revisions that match the specific pattern
                                revisions.Item(i).RevIssue.attRefText = ""
                                revisions.Item(i).RevChangeNote.attRefText = ""
                                revisions.Item(i).RevDate.attRefText = ""
                            End If
                        Next
                        'And then check for RN###### revisions.
                        Dim RNChngNote As Integer = revisions.FindLastIndex(AddressOf FindLastRNChngNote)
                        'Dim RNChngNote As Integer = revisions.FindLastIndex(Function(rev As Revision) rev.RevChangeNote.attRefText Like "RN*")
                        If Not RNChngNote = -1 Then 'ie we found a match
                            For i As Integer = RNChngNote + 1 To revisions.Count - 1
                                revisions.Item(i).RevIssue.attRefText = ""
                                revisions.Item(i).RevChangeNote.attRefText = ""
                                revisions.Item(i).RevDate.attRefText = ""
                            Next
                        ElseIf RNChngNote = -1 Then ' there aren't any RN###### change notes so empty any others.
                            For i As Integer = 0 To revisions.Count - 1
                                revisions.Item(i).RevIssue.attRefText = ""
                                revisions.Item(i).RevChangeNote.attRefText = ""
                                revisions.Item(i).RevDate.attRefText = ""
                            Next
                        End If

                        Dim ndx As Integer = revisions.FindIndex(Function(rev As Revision) IsNumeric(rev.RevIssue.attRefText))
                        Dim numericRevs As Revision = revisions.Find(Function(rev As Revision) IsNumeric(rev.RevIssue.attRefText))
                        If Not numericRevs Is Nothing And Not ndx = 0 Then 'our revisions list has a numeric value within it.
                            numericRevs.RevChangeNote.attRefText = ReleaseNote
                            numericRevs.RevDate.attRefText = ReleaseDate
                            IssueChar = GetNextCode(revisions.Item(ndx - 1).RevIssue.attRefText)
                            numericRevs.RevIssue.attRefText = IssueChar
                        ElseIf Not numericRevs Is Nothing And ndx = 0 Then 'numeric revision found and it IS the first one!
                            revisions.Item(ndx).RevChangeNote.attRefText = ReleaseNote
                            revisions.Item(ndx).RevDate.attRefText = ReleaseDate
                            IssueChar = "A"
                            revisions.Item(ndx).RevIssue.attRefText = IssueChar
                        Else 'no numeric revision found, possibly none at all

                            'then find the first empty revision.
                            ndx = revisions.FindIndex(Function(rev As Revision) rev.RevIssue.attRefText = "") ' should find the first empty issue
                            If ndx = 0 Then ' no revisions on the drawing frame!
                                revisions.Item(ndx).RevChangeNote.attRefText = ReleaseNote
                                revisions.Item(ndx).RevDate.attRefText = ReleaseDate
                                IssueChar = "A"
                                revisions.Item(ndx).RevIssue.attRefText = IssueChar
                            ElseIf ndx = -1 Then 'all rows are used and none of them are numeric revisions.
                                Dim cnt As Integer = revisions.Count
                                For index = 0 To revisions.Count - 1
                                    If Not index = cnt - 1 Then
                                        'swap with next row:
                                        revisions.Item(index).RevChangeNote.attRefText = revisions.Item(index + 1).RevChangeNote.attRefText
                                        revisions.Item(index).RevDate.attRefText = revisions.Item(index + 1).RevDate.attRefText
                                        revisions.Item(index).RevIssue.attRefText = revisions.Item(index + 1).RevIssue.attRefText
                                    Else
                                        revisions.Item(index).RevChangeNote.attRefText = ReleaseNote
                                        revisions.Item(index).RevDate.attRefText = ReleaseDate
                                        IssueChar = GetNextCode(revisions.Item(index - 1).RevIssue.attRefText)
                                        revisions.Item(index).RevIssue.attRefText = IssueChar
                                    End If
                                Next
                            Else 'found the highest empty cell.
                                revisions.Item(ndx).RevChangeNote.attRefText = ReleaseNote
                                revisions.Item(ndx).RevDate.attRefText = ReleaseDate
                                IssueChar = GetNextCode(revisions.Item(ndx - 1).RevIssue.attRefText)
                                revisions.Item(ndx).RevIssue.attRefText = IssueChar
                            End If
                        End If
                        Dim tmplist As List(Of Revision) = revisions.FindAll(Function(rev As Revision) IsNumeric(rev.RevIssue.attRefText))
                        If tmplist.Count > 0 Then
                            For Each rev As Revision In tmplist
                                Dim chngIndex As Integer = revisions.FindIndex(Function(rv As Revision) rv.RevChangeNote.attrefId = rev.RevChangeNote.attrefId)
                                revisions.Item(chngIndex).RevChangeNote.attRefText = ""
                                Dim dateindex As Integer = revisions.FindIndex(Function(rv As Revision) rv.RevDate.attrefId = rev.RevDate.attrefId)
                                revisions.Item(dateindex).RevDate.attRefText = ""
                                Dim issueIndex As Integer = revisions.FindIndex(Function(rv As Revision) rv.RevIssue.attrefId = rev.RevIssue.attrefId)
                                revisions.Item(issueIndex).RevIssue.attRefText = ""
                            Next
                        End If
                        For Each id As ObjectId In attcoll
                            For Each rev As Revision In revisions
                                Dim attref As AttributeReference
                                If rev.RevChangeNote.attrefId = id Then
                                    attref = id.GetObject(OpenMode.ForWrite)
                                    attref.TextString = rev.RevChangeNote.attRefText
                                ElseIf rev.RevDate.attrefId = id Then
                                    attref = id.GetObject(OpenMode.ForWrite)
                                    attref.TextString = rev.RevDate.attRefText
                                ElseIf rev.RevIssue.attrefId = id Then
                                    attref = id.GetObject(OpenMode.ForWrite)
                                    attref.TextString = rev.RevIssue.attRefText
                                End If
                            Next
                        Next
                        'commit changes back to the database
                        tr.Commit()
                    End Using


                    'a bit of cleanup between runs!
                    revisions = Nothing
                    'naming format is: {Drawing-Number}_{sht-###}_{iss-##X}-00.dwg
                    '3DT-957840_sht-001_iss-00b-00.dwg
                    Dim pad As Char = "0"c
                    Dim newfilename As String = drawingpath & dwgnum & "_sht-" & shtnum.PadLeft(3, pad) & "_iss-" & IssueChar.PadLeft(3, pad) & "-00.dwg"
                    Active.Database.SaveAs(newfilename, True, DwgVersion.AC1024, Active.Database.SecurityParameters)
                    Dim serializerd As New XmlSerializer(GetType(Drawings))
                    Dim fsd As New FileStream("C:\Temp\Drawings.xml", FileMode.Open)
                    Dim readerD As XmlReader = XmlReader.Create(fsd)
                    Dim drawingreport As Drawings
                    drawingreport = CType(serializerd.Deserialize(readerD), Drawings)
                    fsd.Close()
                    For Each dwg In drawingreport.Drawing
                        If Path.GetFileNameWithoutExtension(dwg.oldname) = Path.GetFileNameWithoutExtension(originalfilename) Then
                            Dim fn As String = "C:\temp\" & Path.GetFileNameWithoutExtension(newfilename) & "_After.png"
                            Dim pntcoll As Point3dCollection = New Point3dCollection
                            pntcoll.Add(New Point3d(ext.MinPoint.X, ext.MinPoint.Y, ext.MinPoint.Z))
                            pntcoll.Add(New Point3d(ext.MinPoint.X, ext.MaxPoint.Y, ext.MinPoint.Z))
                            pntcoll.Add(New Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, ext.MaxPoint.Z))
                            pntcoll.Add(New Point3d(ext.MaxPoint.X, ext.MinPoint.Y, ext.MinPoint.Z))
                            DeleteMySelection(pntcoll, True)
                            dwg.AfterImgURL = ScreenShotToFile(ext.MinPoint, ext.MaxPoint, fn)
                            dwg.name = dwgnum & "_sht-" & shtnum.PadLeft(3, pad) & "_iss-" & IssueChar.PadLeft(3, pad) & "-00.dwg"
                            dwg.path = drawingpath
                            dwg.revisiondatestr = ReleaseDate
                            dwg.revision = IssueChar
                            Exit For
                        End If
                    Next
                    fsd = New FileStream("C:\Temp\Drawings.xml", FileMode.Create)
                    Dim writer As New XmlTextWriter(fsd, Encoding.Unicode)
                    writer.Formatting = Formatting.Indented
                    serializerd.Serialize(writer, drawingreport)
                    writer.Close()
                    Active.Editor.Regen()
                    Active.Editor.UpdateScreen()
                End Using

                If Not IO.Directory.Exists(supercededPath) Then
                    IO.Directory.CreateDirectory(supercededPath)
                End If
                If Not File.Exists(supercededPath & strdrawingfilename) Then 'should never be true as we're copying the original contents of \CURRENT to \SUPERCEDED before we process anything!
                    File.Move(originalfilename, supercededPath & strdrawingfilename)
                Else 'then delete the one in the current folder:
                    File.Delete(drawingpath & strdrawingfilename)
                End If
            Catch ex As Exception
                Active.WriteMessage(vbLf & "The error was: " & ex.Message)
                revisions = Nothing
            End Try
        End Sub

        ''' <summary>
        ''' Gets the next Alpha character based on the alphaCode input
        ''' </summary>
        ''' <param name="alphaCode">the code to find the next alpha character from</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNextCode(alphaCode As String) As String
            'this debug assert is no longer necessary- we know that everything works except some of the IL drawings.
            'Debug.Assert(alphaCode.Length = 1 AndAlso Regex.IsMatch(alphaCode, "[a-yA-y]"))
            Dim [next] = ChrW(Convert.ToInt32(alphaCode(0)) + 1)
            'Dim [next] = ChrW(alphaCode(0) + 1)
            Return [next].ToString()
        End Function

        ''' <summary>
        ''' Gets the parent directory of the path input based on the parentCount supplied.
        ''' </summary>
        ''' <param name="path">the path we want to look for the parent of.</param>
        ''' <param name="parentCount">the number of folders higher we're looking for.</param>
        ''' <returns>Returns the path as a String</returns>
        ''' <remarks>I'm pretty sure there must be a .NET implementation of this built into System.IO; I just haven't bothered to look for it yet since
        ''' this one works so well.
        ''' </remarks>
        Public Function GetParentDirectory(path As String, parentCount As Integer) As String
            If String.IsNullOrEmpty(path) OrElse parentCount < 1 Then
                Return path
            End If

            Dim parent As String = System.IO.Path.GetDirectoryName(path)

            If System.Threading.Interlocked.Decrement(parentCount) > 0 Then
                Return GetParentDirectory(parent, parentCount)
            End If

            Return parent
        End Function

        ''' <summary>
        ''' Checks whether the Revision passed to it has an empty string value
        ''' </summary>
        ''' <param name="s">The Revision we're checking</param>
        ''' <returns>Returns True or False</returns>
        ''' <remarks></remarks>
        Private Function FirstEmptyChangeNoteString(ByVal s As Revision) As Boolean
            Return s.RevChangeNote.attRefText.Length = 0
        End Function

        ''' <summary>
        ''' Gets the relevant AttributeReference we need based on a regex passed to it.
        ''' </summary>
        ''' <param name="ChangeNoteSearchTerm"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        'Private Function GetAttributeReferenceFromQMA(SearchTerm As Regex, gGroup As Object) As AttRef
        '    Return (From item In gGroup.groupName
        '              Let matches = SearchTerm.Matches(item.b.Tag)
        '              Where (matches.Count > 0)
        '              Select item.b).FirstOrDefault()
        'End Function

        ''' <summary>
        ''' Groups a selection of text/mtext entities together and moves them away from our new border.
        ''' </summary>
        ''' <param name="ext">the area to search.</param>
        ''' <param name="blockId">the ObjectId of our new drawing frame.</param>
        ''' <remarks></remarks>
        Private Sub GroupAndMoveTextNotes(ext As Extents3d, blockId As ObjectId)
            'determine the centerpoint of our view.
            Dim extents As Point3d = New Point3d(
                                    ext.MinPoint.X + ((ext.MaxPoint.X - ext.MinPoint.X) / 2.0), _
                                    ext.MinPoint.Y + ((ext.MaxPoint.Y - ext.MinPoint.Y) / 2.0), _
                                    0)
            Dim pnts As New Point3dCollection
            Dim txtId As ObjectId
            Dim notesinsertpnt As Point3d = Point3d.Origin
            Dim mtxt As MText = New MText
            pnts.Add(ext.MinPoint)
            pnts.Add(extents)
            Dim str As String = CollectTextFromArea(tmppntcoll:=pnts, _
                                                    locName:="NOTES", _
                                                    NotesInsertPnt:=notesinsertpnt, _
                                                     mtxt:=mtxt)
            If Not str = "" Then 'And Not str.StartsWith("CABLES AND COLOURS") Then
                Dim tr As Transaction = Active.Database.TransactionManager.StartTransaction()
                Try
                    '' Open the Block table for read
                    Dim acBlkTbl As BlockTable
                    acBlkTbl = tr.GetObject(Active.Database.BlockTableId, _
                                                 OpenMode.ForRead)

                    '' Open the Block table record current space for write
                    Dim acBlkTblRec As BlockTableRecord
                    acBlkTblRec = tr.GetObject(Active.Database.CurrentSpaceId, _
                                                    OpenMode.ForWrite)

                    '' Create a multiline text object
                    Using acMText As MText = New MText()
                        acMText.Location = notesinsertpnt
                        acMText.TextHeight = mtxt.TextHeight
                        acMText.Contents = str
                        acMText.Layer = "0"
                        acBlkTblRec.AppendEntity(acMText)
                        tr.AddNewlyCreatedDBObject(acMText, True)
                        txtId = acMText.ObjectId
                    End Using
                Catch ex As Exception
                    Active.WriteMessage(vbLf & "The error was: " & ex.Message)
                Finally
                    tr.Commit()
                    tr.Dispose()
                End Try
                ADNIntersectsWith(txtId, blockId)
            End If
        End Sub
#Region "Line and Text Collection/Deletion Methods"


        ''' <summary>
        ''' Deletes a selection based on the Point3dCollection created below.
        ''' </summary>
        ''' <remarks></remarks>
        <CommandMethod("DELSEL")> _
        Public Sub SelectLines()
            'need a set of points to capture existing lines
            Dim p1 As New Point3d(0.0, 0.0, 0.0) 'bottom left
            Dim p2 As New Point3d(0.0, 1000.0, 0.0) 'top left
            Dim p3 As New Point3d(1000.0, 1000.0, 0.0) 'top right
            Dim p4 As New Point3d(1000.0, 0.0, 0.0) 'bottom right

            Dim pntCol As New Point3dCollection()
            pntCol.Add(p1)
            pntCol.Add(p2)
            pntCol.Add(p3)
            pntCol.Add(p4)
            DeleteMySelection(pntCol)
        End Sub

        ''' <summary>
        ''' Deletes a selectionset of lines based on the Point3dCollection passed to it.
        ''' </summary>
        ''' <param name="pntCol">Point3dCollection containing the extents of the area we wish to delete lines from.</param>
        ''' <remarks></remarks>
        Public Sub DeleteMySelection(ByVal pntCol As Point3dCollection, Optional onlyblocks As Boolean = True)
            'need a set of points to capture existing lines

            Dim numOfEntsFound As Integer = 0

            Dim pmtSelRes As PromptSelectionResult = Nothing
            'Dim acTypValAr() As TypedValue = New TypedValue() {New TypedValue(DxfCode.Start, "LINE,INSERT,TEXT,MTEXT,IMAGE")}
            Dim acTypValAr() As TypedValue = New TypedValue() {New TypedValue(DxfCode.Start, "*")}
            Dim selFilter As New SelectionFilter(acTypValAr)
            pmtSelRes = Active.Editor.SelectCrossingPolygon(pntCol, selFilter)
            Using tr As Transaction = Active.Document.Database.TransactionManager.StartTransaction
                Dim bt As BlockTable = DirectCast(tr.GetObject(Active.Database.BlockTableId, OpenMode.ForRead), BlockTable)
                If pmtSelRes.Status = PromptStatus.OK Then
                    For Each objId As ObjectId In pmtSelRes.Value.GetObjectIds()
                        numOfEntsFound += 1
                        'we want to look for and erase blockreferences that aren't our 5.2 border:
                        Dim br As BlockReference = TryCast(objId.GetObject(OpenMode.ForRead), BlockReference)
                        If br IsNot Nothing Then
                            If Not br.Name Like "*5.2(block)" And Not br.Name = "3rd ANGLE + SCALE NOTE" Then
                                br.UpgradeOpen()
                                br.[Erase]()
                            End If
                        Else 'everything else
                            If onlyblocks = False Then
                                Dim ent As Entity = tr.GetObject(objId, OpenMode.ForRead)
                                ent.UpgradeOpen()
                                ent.[Erase]()
                            End If
                        End If
                    Next
                    tr.Commit()
                    Active.WriteMessage(vbCrLf & "Entities erased " & numOfEntsFound.ToString() & vbCrLf)
                Else
                    Active.WriteMessage(vbCrLf & "Did not find entities" & vbCrLf)
                End If
            End Using
        End Sub

        ''' <summary>
        ''' This is the manual version provided by Kean Walmsley...
        ''' </summary>
        ''' <remarks></remarks>
        <CommandMethod("STXT")> _
        Public Sub StrikeTextArea()
            ' Ask for a point. Could also ask for an enclosed text entity
            ' and use the picked point or the text's insertion point

            Dim ppo = New PromptPointOptions(vbLf & "Select point")
            Dim ppr = Active.Editor.GetPoint(ppo)
            If ppr.Status <> PromptStatus.OK Then
                Return
            End If

            StrikeTextArea(ppr.Value)
        End Sub

        ''' <summary>
        ''' The automated version that would work correctly if the EverythingisAWEsome drawing frames had fully enclosed text boxes.
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub StrikeTextArea(ByVal point As Point3d)
            Using tr = Active.Database.TransactionManager.StartTransaction()
                Dim oc = Active.Editor.TraceBoundary(point, False)
                If oc.Count = 1 Then
                    Dim ent = TryCast(oc(0), Entity)
                    If ent IsNot Nothing Then
                        Dim ext = ent.GeometricExtents
                        Dim pt1 = New Point3d(ext.MinPoint.X, ext.MaxPoint.Y, 0)
                        Dim pt2 = New Point3d(ext.MaxPoint.X, ext.MinPoint.Y, 0)
                        ' Create a line between these points
                        Dim ln = New Line(pt1, pt2)
                        ' Add it to the current space
                        Dim btr = DirectCast(tr.GetObject(Active.Database.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                        btr.AppendEntity(ln)
                        tr.AddNewlyCreatedDBObject(ln, True)
                    End If
                End If
                ' Commit even if we didn't add anything
                tr.Commit()
            End Using
        End Sub

        ''' <summary>
        ''' This is the manual CollecTextFromArea version for testing purposes.
        ''' </summary>
        ''' <remarks></remarks>
        <CommandMethod("CTXT")> _
        Public Sub CollectTextFromArea()
            ' Ask for a point. Could also ask for an enclosed text entity
            ' and use the picked point or the text's insertion point

            Dim ppo = New PromptPointOptions(vbLf & "Select point")
            Dim ppr = Active.Editor.GetPoint(ppo)
            If ppr.Status <> PromptStatus.OK Then
                Return
            End If

            Active.WriteMessage(CollectTextFromArea(ppr.Value))
        End Sub

        ''' <summary>
        ''' This is the CollectTextFromArea version that can be called from our other subs.
        ''' </summary>
        ''' <param name="point">Point3d required to make a TraceBoundary attempt from.</param>
        ''' <remarks></remarks>
        Public Function CollectTextFromArea(ByVal point As Point3d, Optional reverseSort As Boolean = False) As String
            Dim tmpstr As String = String.Empty
            Dim numOfEntsFound As Integer = 0
            Dim tmplist As List(Of FloatingText) = New List(Of FloatingText)
            Using tr = Active.Database.TransactionManager.StartTransaction()
                Dim oc = Active.Editor.TraceBoundary(point, False)
                If oc.Count = 1 Then
                    Dim ent = TryCast(oc(0), Entity)
                    If ent IsNot Nothing Then
                        Dim pmtSelRes As PromptSelectionResult = Nothing
                        Dim selFilter As New SelectionFilter(New TypedValue() {New TypedValue(0, "*text")})
                        Dim ext = ent.GeometricExtents
                        Dim pt1 = New Point3d(ext.MinPoint.X, ext.MinPoint.Y, 0) 'bottom left
                        Dim pt2 = New Point3d(ext.MinPoint.X, ext.MaxPoint.Y, 0) 'top left
                        Dim pt3 = New Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, 0) 'top right
                        Dim pt4 = New Point3d(ext.MaxPoint.X, ext.MinPoint.Y, 0) 'bottom right
                        Dim pnts As Point3dCollection = New Point3dCollection()
                        pnts.Add(pt1)
                        pnts.Add(pt2)
                        pnts.Add(pt3)
                        pnts.Add(pt4)
                        pmtSelRes = Active.Editor.SelectCrossingPolygon(pnts, selFilter)

                        If pmtSelRes.Status = PromptStatus.OK Then
                            For Each objId As ObjectId In pmtSelRes.Value.GetObjectIds()
                                numOfEntsFound += 1
                                Dim obj As DBObject = tr.GetObject(objId, OpenMode.ForRead)
                                Dim txt As DBText = TryCast(obj, DBText)
                                If txt IsNot Nothing Then
                                    txt.UpgradeOpen()
                                    txt.[Erase]()
                                    Dim tmpFT As FloatingText = New FloatingText
                                    tmpFT.StringLocation = txt.Position
                                    tmpFT.StringValue = txt.TextString
                                    tmplist.Add(tmpFT)
                                    tmpFT = Nothing
                                Else 'failed converting to DBText
                                    Dim mtxt As MText = TryCast(obj, MText)
                                    If mtxt IsNot Nothing Then
                                        mtxt.UpgradeOpen()
                                        mtxt.[Erase]()
                                    End If
                                    Dim tmpFT As FloatingText = New FloatingText
                                    tmpFT.StringLocation = mtxt.Location
                                    tmpFT.StringValue = mtxt.Contents
                                    tmplist.Add(tmpFT)
                                    tmpFT = Nothing
                                End If
                            Next
                            Dim sorted
                            If reverseSort Then
                                sorted = (From pt As FloatingText In tmplist
                                          Order By pt.StringLocation.Y, pt.StringLocation.X Ascending
                                          Select pt)
                            Else
                                sorted = (From pt As FloatingText In tmplist
                                          Order By pt.StringLocation.Y Descending
                                          Order By pt.StringLocation.X Ascending
                                          Select pt)
                            End If
                            For Each pt As FloatingText In sorted
                                If tmpstr.Length > 0 Then
                                    tmpstr = tmpstr & " " & pt.StringValue
                                Else
                                    tmpstr = pt.StringValue
                                End If
                            Next
                            Active.WriteMessage("MText & Text Entities found " & numOfEntsFound.ToString())
                        Else
                            Active.WriteMessage(vbLf & "Did not find entities")
                        End If
                    End If
                End If

                ' Commit even if we didn't add anything

                tr.Commit()
            End Using
            Return tmpstr
        End Function

        ''' <summary>
        ''' This is a version of CollectTextFromArea that requires a Point3dCollection to collect text from within.
        ''' Deprecated in favour of CollectTextListFromArea
        ''' </summary>
        ''' <param name="tmppntcoll">Point3dCollection to check for text within.</param>
        ''' <param name="reverseSort"></param>
        ''' <returns>String value found with tmppntcoll boundary</returns>
        ''' <remarks></remarks>
        Public Function CollectTextFromArea(tmppntcoll As Point3dCollection, _
                                            locName As String, _
                                            ByRef NotesInsertPnt As Point3d, _
                                            ByRef mtxt As MText, _
                                            Optional reverseSort As Boolean = False, _
                                            Optional ByRef lineCount As Integer = 0 _
                                            ) As String
            Dim tmpstr As String = String.Empty
            'Dim pts As Point3dCollection = New Point3dCollection()
            Dim numOfEntsFound As Integer = 0
            Dim tmplist As List(Of FloatingText) = New List(Of FloatingText)
            Using tr = Active.Database.TransactionManager.StartTransaction()
                Dim pmtSelRes As PromptSelectionResult = Nothing
                Dim acTypValAr() As TypedValue = New TypedValue() {
                    New TypedValue(DxfCode.Start, "ATTDEF,TEXT,MTEXT")}

                ' in case you want to see what the filter looks like
                'For Each tv As TypedValue In acTypValAr
                '   Active.WriteMessage(String.Format("\nFilter Pair: {0}", tv.ToString()))
                'Next

                Dim selFilter As New SelectionFilter(acTypValAr)
                If locName = "NOTES" Then
                    pmtSelRes = Active.Editor.SelectCrossingWindow(tmppntcoll.Item(0), tmppntcoll.Item(1), selFilter)
                Else
                    pmtSelRes = Active.Editor.SelectCrossingPolygon(tmppntcoll, selFilter)
                End If

                If pmtSelRes.Status = PromptStatus.OK Then
                    For Each objId As ObjectId In pmtSelRes.Value.GetObjectIds()
                        Dim obj As DBObject = tr.GetObject(objId, OpenMode.ForRead)
                        Dim attdef As AttributeDefinition = TryCast(obj, AttributeDefinition)
                        If attdef IsNot Nothing Then
                            numOfEntsFound += 1
                            attdef.UpgradeOpen()
                            attdef.[Erase]()
                            Dim tmpFT As FloatingText = New FloatingText
                            tmpFT.StringLocation = attdef.Position
                            tmpFT.TextHeight = attdef.Height
                            If attdef.TextString = "" Then
                                tmpFT.StringValue = attdef.Tag
                            Else
                                tmpFT.StringValue = attdef.TextString
                            End If
                            tmplist.Add(tmpFT)
                            tmpFT = Nothing
                        Else
                            Dim txt As DBText = TryCast(obj, DBText)
                            If txt IsNot Nothing Then
                                If locName.ToUpper() = "NOTES" Then
                                    If Matchesdisallowednotes(txt.TextString) Then
                                        Continue For 'basically we need to skip this entity.
                                    End If
                                End If
                                'If txt.TextString.StartsWith("SEE NOTE") Then
                                '    Continue For
                                'End If
                                'If txt.TextString.StartsWith("CABLES") Then Return "" 'it looked like things were going to kick off so I got outta there!
                                numOfEntsFound += 1
                                txt.UpgradeOpen()
                                txt.[Erase]()
                                Dim tmpFT As FloatingText = New FloatingText
                                tmpFT.StringLocation = txt.Position
                                tmpFT.StringValue = txt.TextString
                                tmpFT.TextHeight = txt.Height
                                tmplist.Add(tmpFT)
                                tmpFT = Nothing
                            Else 'failed converting to DBText
                                Dim mtext As MText = TryCast(obj, MText)
                                If mtext IsNot Nothing Then
                                    If locName.ToUpper() = "NOTES" Then
                                        If Matchesdisallowednotes(mtext.Contents) Then
                                            Continue For 'basically we need to skip this entity.
                                        End If
                                    End If
                                    numOfEntsFound += 1
                                    mtext.UpgradeOpen()
                                    mtext.[Erase]()
                                    Dim tmpFT As FloatingText = New FloatingText
                                    tmpFT.StringLocation = mtext.Location
                                    tmpFT.StringValue = mtext.Contents
                                    tmpFT.TextHeight = mtext.TextHeight
                                    tmplist.Add(tmpFT)
                                    tmpFT = Nothing
                                End If
                            End If
                            End If
                    Next
                    Dim sorted As IEnumerable(Of FloatingText)
                    If reverseSort Then
                        sorted = (From pt As FloatingText In tmplist
                                  Order By pt.StringLocation.Y, pt.StringLocation.X Ascending
                                  Select pt)
                    Else 'removed this line from the following because it caused attributes to get incorrectly sorted: Order By pt.StringLocation.X Ascending
                        sorted = (From pt As FloatingText In tmplist
                                  Order By pt.StringLocation.Y Descending
                                  Select pt)
                    End If

                    Dim dblstr As String = String.Empty
                    Dim skipnextcrlf As Boolean = False
                    For Each pt As FloatingText In sorted
                        If tmpstr.Length > 0 Then
                            If locName.ToUpper() = "TOLERANCES" Then
                                lineCount = lineCount + 1
                                'Tolerance strings seem to be split better than the other long strings we're seeing.
                                'hence we don't necessarily need to checklengthandsplit them.
                                tmpstr = tmpstr & vbCrLf & pt.StringValue
                                'replace +/- with ±:
                                tmpstr = tmpstr.Replace("+/-", "±")
                            ElseIf locName.ToUpper() = "MATERIAL" Then 'Or locName.ToUpper() = "SURFACE_TEXTURE" Or locName.ToUpper() = "PROTECTIVE_FINISH" Then
                                tmpstr = checkLengthAndSplit(tmpstr)
                            ElseIf locName.ToUpper() = "SURFACE_TEXTURE" Or locName.ToUpper() = "PROTECTIVE_FINISH" Then
                                tmpstr = tmpstr & " " & pt.StringValue
                                If MatchesSurfaceTextureFormat(tmpstr) Then
                                    tmpstr = returnReformattedstr(tmpstr, dblstr)
                                End If
                                'needed to move this here because string.split() doesn't like it if there are VBcrlf characters in the string.
                                tmpstr = checkLengthAndSplit(tmpstr)
                            ElseIf locName.ToUpper() = "NOTES" Then
                                mtxt.Location = pt.StringLocation
                                mtxt.TextHeight = pt.TextHeight
                                If MatchesNoteBulletPoint(pt.StringValue) Then
                                    skipnextcrlf = True
                                    tmpstr = tmpstr & vbCrLf & pt.StringValue
                                ElseIf Not MatchesNoteBulletPoint(pt.StringValue) And skipnextcrlf Then
                                    skipnextcrlf = False
                                    tmpstr = tmpstr & " " & pt.StringValue
                                Else
                                    tmpstr = tmpstr & vbCrLf & pt.StringValue
                                End If

                                'tmpstr = checkLengthAndSplit(tmpstr, True)
                            Else
                                tmpstr = tmpstr & " " & pt.StringValue
                                tmpstr = checkLengthAndSplit(tmpstr)
                            End If
                        Else
                            If locName.ToUpper() = "MATERIAL" Then
                                tmpstr = checkLengthAndSplit(pt.StringValue)
                            ElseIf locName.ToUpper() = "SURFACE_TEXTURE" Or locName.ToUpper() = "PROTECTIVE_FINISH" Then
                                tmpstr = checkLengthAndSplit(pt.StringValue)
                                Dim number As Double
                                If Double.TryParse(tmpstr, number) Then 'tmpstr is probably something like 0.8 or similar.
                                    dblstr = tmpstr & "µm"
                                End If
                            ElseIf locName.ToUpper() = "NOTES" Then
                                If NotesInsertPnt = Point3d.Origin Then
                                    NotesInsertPnt = pt.StringLocation
                                    'NotesInsertPnt = tmplist.MaxBy(Function(ft As FloatingText) ft.StringLocation.Y >= ft.StringLocation.Y).StringLocation
                                End If
                                mtxt.Location = pt.StringLocation
                                mtxt.TextHeight = pt.TextHeight
                                tmpstr = pt.StringValue
                                'tmpstr = checkLengthAndSplit(pt.StringValue, True)
                            Else
                                tmpstr = pt.StringValue
                            End If
                        End If
                    Next
                    Active.WriteMessage(vbLf & "MText & Text Entities found " & numOfEntsFound.ToString())
                Else
                    Active.WriteMessage(vbLf & "Did not find entities")
                End If

                ' Commit even if we didn't add anything

                tr.Commit()
            End Using
            Return tmpstr
        End Function

        ''' <summary>
        ''' This overload allows us to insert dummy objects into our main CollectTextFromArea Method.
        ''' </summary>
        ''' <param name="tmppntcoll"></param>
        ''' <param name="locName"></param>
        ''' <param name="reverseSort"></param>
        ''' <param name="lineCount"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CollectTextFromArea(tmppntcoll As Point3dCollection, locName As String, reverseSort As Boolean, ByRef lineCount As Integer) As String
            Dim pnt As Point3d = Point3d.Origin
            Dim mtxt As MText = New MText
            Return CollectTextFromArea(tmppntcoll:=tmppntcoll, locName:=locName, NotesInsertPnt:=pnt, lineCount:=lineCount, reverseSort:=reverseSort, mtxt:=mtxt)
        End Function

        ''' <summary>
        ''' returns our tmpstr reformatted correctly
        ''' </summary>
        ''' <param name="tmpstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function returnReformattedstr(tmpstr As String, dblstr As String) As String
            Dim splitstr As New List(Of String)(tmpstr.Split(" "))
            Dim machinestr As String = String.Empty
            For j As Integer = 0 To splitstr.FindIndex(Function(str As String) str = "") - 1
                Dim number As Integer
                If Not Double.TryParse(splitstr.Item(j), number) Then
                    If machinestr = "" Then
                        machinestr = splitstr.Item(j)
                    Else
                        machinestr = machinestr & " " & splitstr.Item(j)
                    End If
                End If
            Next
            Dim spaceint As Integer = splitstr.FindLastIndex(Function(rev As String) rev = "")
            Dim rebuiltstr As String = String.Empty
            For i = spaceint + 1 To splitstr.Count - 1
                If splitstr.Item(i).StartsWith(" ") Then
                    rebuiltstr = rebuiltstr & splitstr.Item(i)
                Else
                    If rebuiltstr = "" Then
                        rebuiltstr = splitstr.Item(i)
                    Else
                        rebuiltstr = rebuiltstr & " " & splitstr.Item(i)
                    End If
                End If
            Next
            tmpstr = machinestr & " " & dblstr & " " & rebuiltstr
            Return tmpstr.Replace("  ", " ")
        End Function

        ''' <summary>
        ''' check the length of a string and split if necessary.
        ''' </summary>
        ''' <param name="tmpstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function checkLengthAndSplit(tmpstr As String, Optional isNotes As Boolean = False) As String
            Dim lengthint As Integer = 0
            If isNotes = True Then
                lengthint = 100
            Else
                lengthint = 30
            End If

            If tmpstr.Length > lengthint Then
                Dim part1 As String = tmpstr.Substring(0, CInt(tmpstr.Length / 2))
                Dim part2 As String = tmpstr.Substring(CInt(tmpstr.Length / 2))
                Dim p1 As Integer = part1.LastIndexOf(" ")
                Dim p2 As Integer = part2.IndexOf(" ")
                If tmpstr.Length - p1 = 0 OrElse p2 = 0 Then
                    Return part1 & vbCrLf & part2
                ElseIf p1 <> -1 AndAlso tmpstr.Length - p1 < p2 Then
                    Return part1.Substring(0, p1) & vbCrLf & part1.Substring(p1) & part2
                ElseIf p2 <> -1 AndAlso tmpstr.Length - p1 > p2 Then
                    Return part1 & part2.Substring(0, p2) & vbCrLf & part2.Substring(p2)
                Else
                    Return tmpstr
                End If
            Else
                Return tmpstr
            End If
        End Function

        ''' <summary>
        ''' This is a version that will accept a Point3dCollection
        ''' </summary>
        ''' <param name="tmppntcoll">Point3dCollection where we wish to search for text</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CollectTextListFromArea(tmppntcoll As Point3dCollection, ByVal attstr As String) As List(Of String)
            Dim tmpstrlist As List(Of String) = New List(Of String)
            'Dim pts As Point3dCollection = New Point3dCollection()
            Dim numOfEntsFound As Integer = 0
            Dim tmplist As List(Of FloatingText) = New List(Of FloatingText)
            Dim tmpstr As String = String.Empty
            Using tr = Active.Database.TransactionManager.StartTransaction()
                Dim pmtSelRes As PromptSelectionResult = Nothing
                Dim acTypValAr(0) As TypedValue
                acTypValAr.SetValue(New TypedValue(DxfCode.Start, "*"), 0)
                Dim selFilter As New SelectionFilter(acTypValAr)
                pmtSelRes = Active.Editor.SelectCrossingPolygon(tmppntcoll, selFilter)
                If pmtSelRes.Status = PromptStatus.OK Then
                    For Each objId As ObjectId In pmtSelRes.Value.GetObjectIds()
                        Dim obj As DBObject = tr.GetObject(objId, OpenMode.ForRead)
                        Dim attdef As AttributeDefinition = TryCast(obj, AttributeDefinition)
                        If attdef IsNot Nothing Then
                            numOfEntsFound += 1
                            attdef.UpgradeOpen()
                            attdef.[Erase]()
                            Dim tmpFT As FloatingText = New FloatingText
                            tmpFT.StringLocation = attdef.Position
                            If attdef.TextString = "" Then
                                tmpFT.StringValue = attdef.Tag
                            Else
                                tmpFT.StringValue = attdef.TextString
                            End If
                            tmplist.Add(tmpFT)
                            tmpFT = Nothing
                        Else
                            Dim txt As DBText = TryCast(obj, DBText)
                            If txt IsNot Nothing Then
                                numOfEntsFound += 1
                                txt.UpgradeOpen()
                                txt.[Erase]()
                                Dim tmpFT As FloatingText = New FloatingText
                                tmpFT.StringLocation = txt.Position
                                tmpFT.StringValue = txt.TextString
                                tmplist.Add(tmpFT)
                                tmpFT = Nothing
                            Else 'failed converting to DBText
                                Dim mtxt As MText = TryCast(obj, MText)
                                If mtxt IsNot Nothing Then
                                    numOfEntsFound += 1
                                    mtxt.UpgradeOpen()
                                    mtxt.[Erase]()
                                    Dim tmpFT As FloatingText = New FloatingText
                                    tmpFT.StringLocation = mtxt.Location
                                    tmpFT.StringValue = mtxt.Contents
                                    tmplist.Add(tmpFT)
                                    tmpFT = Nothing
                                End If
                            End If
                        End If
                    Next
                    Dim sorted = (From pt As FloatingText In tmplist
                                  Order By pt.StringLocation.Y Descending
                                  Select pt)

                    For Each pt As FloatingText In sorted
                        Dim tmpint As Integer = InStrRev(pt.StringValue, "/")
                        If tmpstr.Length > 0 Then
                            tmpint = InStrRev(pt.StringValue, "/")
                            If tmpint = 0 Then 'the last character is a slash
                                tmpstr = tmpstr & pt.StringValue
                                tmpstrlist.Add(tmpstr)
                                tmpstr = String.Empty
                            Else
                                tmpstr = tmpstr & " " & pt.StringValue
                            End If
                        ElseIf Not tmpint = pt.StringValue.Length Then
                            tmpstr = pt.StringValue
                            tmpstrlist.Add(tmpstr)
                            tmpstr = String.Empty
                        ElseIf tmpint = pt.StringValue.Length Then 'last character is /
                            tmpstr = pt.StringValue
                            'If Regex.IsMatch(tmpstr, "\w{2}\D\d\D\w.*") Then ' matches HR/#/###### format
                            '    tmpstrlist.Add(tmpstr)
                            'End If
                        End If
                    Next
                    Active.WriteMessage(vbCrLf & "MText & Text Entities found " & numOfEntsFound.ToString())
                Else
                    Active.WriteMessage(vbLf & "Did not find entities")
                End If

                ' Commit even if we didn't add anything

                tr.Commit()
            End Using
            Return tmpstrlist
        End Function
#End Region
#Region "Regex Methods"
        ''' <summary>
        ''' Checks the note string against an xml list of disallowed values.
        ''' </summary>
        ''' <param name="p1"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function Matchesdisallowednotes(p1 As String) As Boolean
            Dim asmpath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim serializer As New XmlSerializer(GetType(notes))
            Dim fs As New FileStream(asmpath + "\Resources\SkipNotes.xml", FileMode.Open)
            Dim reader As XmlReader = XmlReader.Create(fs)
            'new these objects here.
            Dim skn As notes = CType(serializer.Deserialize(reader), notes)
            fs.Close()
            For Each note As notesText In skn.text
                If p1.Contains(note.string) Then Return True
            Next

        End Function
        ''' <summary>
        ''' Checks whether the notes match the #. pattern
        ''' </summary>
        ''' <param name="p1"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function MatchesNoteBulletPoint(p1 As String) As Boolean
            Dim pattern As String = "\b\d{1,}\.\z"
            Dim NoteRegex As New Regex(pattern)
            Return NoteRegex.IsMatch(p1)
        End Function
        ''' <summary>
        ''' Checks for whether the revision is of the format 1P1
        ''' </summary>
        ''' <param name="p1">the string to check</param>
        ''' <returns>True or False depending on the outcome of the check.</returns>
        ''' <remarks></remarks>
        Private Function MatchesOddRevision(p1 As String) As Boolean
            Dim pattern As String = "\d[A-Z]\d"
            Dim oddRegex As New Regex(pattern)
            Return oddRegex.IsMatch(p1)
        End Function

        ''' <summary>
        ''' Checks for the last RN Number
        ''' </summary>
        ''' <param name="chngnote">the Revision to check the Change note of.</param>
        ''' <returns>True or False depending on the outcome of the check.</returns>
        ''' <remarks></remarks>
        Private Function FindLastRNChngNote(ByVal chngnote As Revision) As Boolean
            Dim pattern As String = "RN\d{6}" 'exactly matches our RN###### pattern
            'Dim pattern As String = "[A-Z]{2}\d{6}" 'matches our RN###### pattern
            Dim lastRNRegex As New Regex(pattern)
            Return lastRNRegex.IsMatch(chngnote.RevChangeNote.attRefText)

        End Function

        ''' <summary>
        ''' Compares the Surface Texture string against our pattern for it.
        ''' </summary>
        ''' <param name="tmpstr">the string to check</param>
        ''' <returns>True if the string matches our pattern</returns>
        ''' <remarks></remarks>
        Private Function MatchesSurfaceTextureFormat(tmpstr As String) As Boolean
            Dim pattern As String = ".*\s{2,}.*"
            Dim surfacetexture As New Regex(pattern)
            Return surfacetexture.IsMatch(tmpstr)
        End Function

        ''' <summary>
        ''' Checks whether the input string contains any non-alpha-characters
        ''' </summary>
        ''' <param name="inputString">The string to check</param>
        ''' <returns>True or False depending on the outcome of the check.</returns>
        ''' <remarks></remarks>
        Public Function ContainsStars(inputString As String) As Boolean
            Dim r As New Regex(".*\*.*")
            If r.IsMatch(inputString) Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' Gets the C##### number from the filename in the event that the drawing number currently
        ''' contains "*" or "C" or is completely empty.
        ''' </summary>
        ''' <param name="inputstring"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCNumberFromFileName(inputstring As String) As String
            Dim r As New Regex("\w\d{5}")
            Return r.Match(inputstring).Captures.Item(0).ToString()
        End Function

        ''' <summary>
        ''' Checks whether the drawing number is just "C"
        ''' </summary>
        ''' <param name="tmpstr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ContainsJustC(InputString As String) As Boolean
            Dim r As New Regex("C")
            If r.IsMatch(InputString) Then
                Return True
            Else
                Return False
            End If
        End Function
#End Region 'regex methods
        ''' <summary>
        ''' Corrects the drawing number dependant on what blockname the attributereference is from.
        ''' </summary>
        ''' <param name="myAtt">the attributereference to check</param>
        ''' <param name="tag">the tag of the attributereference to check</param>
        ''' <param name="blockname">the block name to filter for</param>
        ''' <returns>a corrected string containing the updated Drawing Number</returns>
        ''' <remarks></remarks>
        Private Function getCorrectedDrawingNumber(mydoc As Document, myAtt As AttributeReference, tag As String, blockname As String) As String
            Dim tmpstr As String = myAtt.TextString
            If blockname Like "Border*" Then
                If tag = "DRAWING NUMBER 1" Or tag = "DRAWING NUMBER 2" Then
                    If ContainsStars(tmpstr) Then 'there's something weird about the drawing number
                        tmpstr = GetCNumberFromFileName(Path.GetFileNameWithoutExtension(mydoc.Database.Filename))
                    ElseIf ContainsJustC(tmpstr) And tmpstr.Length = 1 Then
                        tmpstr = GetCNumberFromFileName(Path.GetFileNameWithoutExtension(mydoc.Database.Filename))
                    ElseIf tmpstr.Length = 0 Then
                        tmpstr = GetCNumberFromFileName(Path.GetFileNameWithoutExtension(mydoc.Database.Filename))
                    End If
                    If blockname = "Border" Then
                        tmpstr = "HR/3/" & tmpstr
                    ElseIf blockname = "Border A0" Then
                        tmpstr = "HR/0/" & tmpstr
                    ElseIf blockname = "Border A1" Then
                        tmpstr = "HR/1/" & tmpstr
                    ElseIf blockname = "Border A2" Then
                        tmpstr = "HR/2/" & tmpstr
                    ElseIf blockname = "Border A3" Then
                        tmpstr = "HR/3/" & tmpstr
                    ElseIf blockname = "Border A4" Then
                        tmpstr = "HR/4/" & tmpstr
                    End If
                Else
                    Return tmpstr
                    Exit Function
                End If
            ElseIf blockname = "A3BORD" Then
                If tag = "DRAWING NUMBER 1" Or tag = "DRAWING NUMBER 2" Then
                    tmpstr = "HR/3/" & tmpstr
                End If
            End If
            Return tmpstr
        End Function



#Region "Kean's Sort a point2dcollection or point3dcollection code"
        Public Sub PrintPoints(ed As Editor, pts As IEnumerable(Of Point3d))
            For Each pt As Point3d In pts
                Active.WriteMessage("{0}" & vbLf, pt)
            Next
        End Sub

        <CommandMethod("PTS")> _
        Public Sub PointSort()
            ' We'll use a random number generator

            Dim rnd As New Random()

            ' Generate the random set of points
            Dim pts = (From n In Enumerable.Range(1, 20)
                       Select New Point3d((2 * rnd.NextDouble() - 1) * 1000, _
                                          (2 * rnd.NextDouble() - 1) * 1000, _
                                          (2 * rnd.NextDouble() - 1) * 1000))
            'Dim pts = From n In Enumerable.Range(1, 20)New Point3d((2 * rnd.NextDouble() - 1) * 1000, (2 * rnd.NextDouble() - 1) * 1000, (2 * rnd.NextDouble() - 1) * 1000)

            ' Order them by X then Y then Z

            Dim sorted = From pt In pts
                         Order By pt.X, pt.Y, pt.Z
                         Select pt

            Active.WriteMessage(vbLf & "Points before sort:" & vbLf & vbLf)
            PrintPoints(Active.Editor, pts)
            Active.WriteMessage(vbLf & vbLf & "Points after sort:" & vbLf & vbLf)
            PrintPoints(Active.Editor, sorted)
        End Sub

        <CommandMethod("CG")> _
        Public Sub CompleteGraph()
            ' Generate points, put them in a list

            Dim alpha As Double = Math.PI * 2 / 19
            Dim pts = (From n In Enumerable.Range(0, 19)
                       Select New Point3d(Math.Cos(n * alpha) * 5, Math.Sin(n * alpha) * 5, 0)).ToList()

            ' Generate lines from points

            Dim lns = From a In pts
                      From b In pts
                      Where pts.IndexOf(b) > pts.IndexOf(a)
                      Select New Line(a, b)

            ' Add them to the current space in the active drawing

            Dim tr As Transaction = Active.Database.TransactionManager.StartTransaction()
            Using tr
                Dim btr As BlockTableRecord = DirectCast(tr.GetObject(Active.Database.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                For Each ln As Line In lns
                    btr.AppendEntity(ln)
                    tr.AddNewlyCreatedDBObject(ln, True)
                Next
                tr.Commit()
            End Using
        End Sub
#End Region

#Region "Kean's Modelspace snapshot code"

        Public Structure AcColorSettings
            'UInt32
            Public dwGfxModelBkColor As UInt32
            Public dwGfxLayoutBkColor As UInt32
            Public dwParallelBkColor As UInt32
            Public dwBEditBkColor As UInt32
            Public dwCmdLineBkColor As UInt32
            Public dwPlotPrevBkColor As UInt32
            Public dwSkyGradientZenithColor As UInt32
            Public dwSkyGradientHorizonColor As UInt32
            Public dwGroundGradientOriginColor As UInt32
            Public dwGroundGradientHorizonColor As UInt32
            Public dwEarthGradientAzimuthColor As UInt32
            Public dwEarthGradientHorizonColor As UInt32
            Public dwModelCrossHairColor As UInt32
            Public dwLayoutCrossHairColor As UInt32
            Public dwParallelCrossHairColor As UInt32
            Public dwPerspectiveCrossHairColor As UInt32
            Public dwBEditCrossHairColor As UInt32
            Public dwParallelGridMajorLines As UInt32
            Public dwPerspectiveGridMajorLines As UInt32
            Public dwParallelGridMinorLines As UInt32
            Public dwPerspectiveGridMinorLines As UInt32
            Public dwParallelGridAxisLines As UInt32
            Public dwPerspectiveGridAxisLines As UInt32
            Public dwTextForeColor As UInt32, dwTextBkColor As UInt32
            Public dwCmdLineForeColor As UInt32
            Public dwCmdLineTempPromptBkColor As UInt32
            Public dwCmdLineTempPromptTextColor As UInt32
            Public dwCmdLineCmdOptKeywordColor As UInt32
            Public dwCmdLineCmdOptBkColor As UInt32
            Public dwCmdLineCmdOptHighlightedColor As UInt32
            Public dwAutoTrackingVecColor As UInt32
            Public dwLayoutATrackVecColor As UInt32
            Public dwParallelATrackVecColor As UInt32
            Public dwPerspectiveATrackVecColor As UInt32
            Public dwBEditATrackVecColor As UInt32
            Public dwModelASnapMarkerColor As UInt32
            Public dwLayoutASnapMarkerColor As UInt32
            Public dwParallelASnapMarkerColor As UInt32
            Public dwPerspectiveASnapMarkerColor As UInt32
            Public dwBEditASnapMarkerColor As UInt32
            Public dwModelDftingTooltipColor As UInt32
            Public dwLayoutDftingTooltipColor As UInt32
            Public dwParallelDftingTooltipColor As UInt32
            Public dwPerspectiveDftingTooltipColor As UInt32
            Public dwBEditDftingTooltipColor As UInt32
            Public dwModelDftingTooltipBkColor As UInt32
            Public dwLayoutDftingTooltipBkColor As UInt32
            Public dwParallelDftingTooltipBkColor As UInt32
            Public dwPerspectiveDftingTooltipBkColor As UInt32
            Public dwBEditDftingTooltipBkColor As UInt32
            Public dwModelLightGlyphs As UInt32
            Public dwLayoutLightGlyphs As UInt32
            Public dwParallelLightGlyphs As UInt32
            Public dwPerspectiveLightGlyphs As UInt32
            Public dwBEditLightGlyphs As UInt32
            Public dwModelLightHotspot As UInt32
            Public dwLayoutLightHotspot As UInt32
            Public dwParallelLightHotspot As UInt32
            Public dwPerspectiveLightHotspot As UInt32
            Public dwBEditLightHotspot As UInt32
            Public dwModelLightFalloff As UInt32
            Public dwLayoutLightFalloff As UInt32
            Public dwParallelLightFalloff As UInt32
            Public dwPerspectiveLightFalloff As UInt32
            Public dwBEditLightFalloff As UInt32
            Public dwModelLightStartLimit As UInt32
            Public dwLayoutLightStartLimit As UInt32
            Public dwParallelLightStartLimit As UInt32
            Public dwPerspectiveLightStartLimit As UInt32
            Public dwBEditLightStartLimit As UInt32
            Public dwModelLightEndLimit As UInt32
            Public dwLayoutLightEndLimit As UInt32
            Public dwParallelLightEndLimit As UInt32
            Public dwPerspectiveLightEndLimit As UInt32
            Public dwBEditLightEndLimit As UInt32
            Public dwModelCameraGlyphs As UInt32
            Public dwLayoutCameraGlyphs As UInt32
            Public dwParallelCameraGlyphs As UInt32
            Public dwPerspectiveCameraGlyphs As UInt32
            Public dwModelCameraFrustrum As UInt32
            Public dwLayoutCameraFrustrum As UInt32
            Public dwParallelCameraFrustrum As UInt32
            Public dwPerspectiveCameraFrustrum As UInt32
            Public dwModelCameraClipping As UInt32
            Public dwLayoutCameraClipping As UInt32
            Public dwParallelCameraClipping As UInt32
            Public dwPerspectiveCameraClipping As UInt32
            Public nModelCrosshairUseTintXYZ As Integer
            Public nLayoutCrosshairUseTintXYZ As Integer
            Public nParallelCrosshairUseTintXYZ As Integer
            Public nPerspectiveCrosshairUseTintXYZ As Integer
            Public nBEditCrossHairUseTintXYZ As Integer
            Public nModelATrackVecUseTintXYZ As Integer
            Public nLayoutATrackVecUseTintXYZ As Integer
            Public nParallelATrackVecUseTintXYZ As Integer
            Public nPerspectiveATrackVecUseTintXYZ As Integer
            Public nBEditATrackVecUseTintXYZ As Integer
            Public nModelDftingTooltipBkUseTintXYZ As Integer
            Public nLayoutDftingTooltipBkUseTintXYZ As Integer
            Public nParallelDftingTooltipBkUseTintXYZ As Integer
            Public nPerspectiveDftingTooltipBkUseTintXYZ As Integer
            Public nBEditDftingTooltipBkUseTintXYZ As Integer
            Public nParallelGridMajorLineTintXYZ As Integer
            Public nPerspectiveGridMajorLineTintXYZ As Integer
            Public nParallelGridMinorLineTintXYZ As Integer
            Public nPerspectiveGridMinorLineTintXYZ As Integer
            Public nParallelGridAxisLineTintXYZ As Integer
            Public nPerspectiveGridAxisLineTintXYZ As Integer
        End Structure

        ' For the coordinate tranformation we need...  

        ' A Win32 function:

        <DllImport("user32.dll")> _
        Private Shared Function ClientToScreen(hWnd As IntPtr, ByRef pt As Point) As Boolean
        End Function

        ' And to access the colours in AutoCAD, we need ObjectARX...

        ' 64 bit, AutoCAD 2013
        <DllImport("accore.dll", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="?acedGetCurrentColors@@YAHPEAUAcColorSettings@@@Z")> _
        Private Shared Function acedGetCurrentColors64(ByRef colorSettings As AcColorSettings) As Boolean
        End Function

        ' 64 bit, AutoCAD 2013
        <DllImport("accore.dll", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="?acedSetCurrentColors@@YAHPEAUAcColorSettings@@@Z")> _
        Private Shared Function acedSetCurrentColors64(ByRef colorSettings As AcColorSettings) As Boolean
        End Function

        ' Helper functions that call automatically to 32- or 64-bit
        ' versions, as appropriate

        Private Shared Function acedGetCurrentColors(ByRef colorSettings As AcColorSettings) As Boolean
            If IntPtr.Size > 4 Then
                Return acedGetCurrentColors64(colorSettings)
            End If
        End Function

        Private Shared Function acedSetCurrentColors(ByRef colorSettings As AcColorSettings) As Boolean
            If IntPtr.Size > 4 Then
                Return acedSetCurrentColors64(colorSettings)
            End If
        End Function
        ''' <summary>
        ''' Command to capture the main and active drawing windows or a user-selected portion of a drawing
        ''' </summary>
        ''' <param name="pnt1"></param>
        ''' <param name="pnt2"></param>
        ''' <param name="filename"></param>
        ''' <remarks></remarks>
        '<CommandMethod("ADNPLUGINS", "SCREENSHOT", CommandFlags.Modal)> _
        Public Shared Function CaptureSnapShot(ByVal pnt1 As Point3d, ByVal pnt2 As Point3d, ByVal filename As String) As String
            ' Retrieve our application settings (or create new ones)

            Dim ad As New AppData()
            ad.Reload()

            If ad IsNot Nothing Then
                'always set our background to white!
                ad.WhiteBackground = True
                ad.Save()

                Dim first As Point3d = pnt1
                Dim second As Point3d = pnt2

                ' Generate screen coordinate points based on the
                ' drawing points selected

                Dim pt1 As Point, pt2 As Point

                ' First we get the viewport number

                Dim vp As Short = CShort(Application.GetSystemVariable("CVPORT"))

                ' Then the handle to the current drawing window

                Dim hWnd As IntPtr = Active.Document.Window.Handle

                ' Now calculate the selected corners in screen coordinates

                pt1 = ScreenFromDrawingPoint(Active.Editor, hWnd, first, vp, True)
                pt2 = ScreenFromDrawingPoint(Active.Editor, hWnd, second, vp, True)

                ' Now save this portion of our screen as a raster image
                Return ScreenShotToFile(pt1, pt2, filename, ad)
            End If
            Return filename
        End Function

        Private Shared Function ScreenFromDrawingPoint(ed As Editor, hWnd As IntPtr, pt As Point3d, vpNum As Short, useUcs As Boolean) As Point
            ' Transform from UCS to WCS, if needed

            Dim wcsPt As Point3d = (If(useUcs, pt.TransformBy(ed.CurrentUserCoordinateSystem), pt))

            Dim winPt As System.Windows.Point = Active.Editor.PointToScreen(wcsPt, vpNum)
            Dim s As System.Windows.Vector = Autodesk.AutoCAD.Windows.Window.GetDeviceIndependentScale(IntPtr.Zero)
            Dim res As New Point(CInt(winPt.X * s.X), CInt(winPt.Y * s.Y))
            ClientToScreen(hWnd, res)
            Return res
        End Function



        Private Shared Function ScreenShotToFile(pt1 As Point, pt2 As Point, filename As String, ad As AppData) As String
            ' Create the top left corner from the two corners
            ' provided (by taking the min of both X and Y values)

            Dim pt As New Point(Math.Min(pt1.X, pt2.X), Math.Min(pt1.Y, pt2.Y))

            ' Determine the size by subtracting X & Y values and
            ' taking the absolute value of each

            Dim sz As New Size(Math.Abs(pt1.X - pt2.X), Math.Abs(pt1.Y - pt2.Y))
            Dim tr As Transaction = Active.Database.TransactionManager.StartTransaction()
            Using tr
                Dim ocs As New AcColorSettings()

                If ad.WhiteBackground Then
                    ' Get the current system colours

                    acedGetCurrentColors(ocs)

                    ' Take a copy - we'll leave the original to reset
                    ' the values later on, once we've finished

                    Dim cs As AcColorSettings = ocs

                    ' Make both background colours white (the 3D
                    ' background isn't currently being picked up)

                    cs.dwGfxModelBkColor = 16777215
                    cs.dwGfxLayoutBkColor = 16777215
                    ' Set the modified colours
                    acedSetCurrentColors(cs)
                    Active.WriteMessage(vbLf)
                    Active.Editor.Regen()
                    Active.Editor.UpdateScreen()
                End If

                ' Set the bitmap object to the size of the window

                Dim bmp As New Bitmap(sz.Width, sz.Height, PixelFormat.Format32bppArgb)
                Using bmp
                    ' Create a graphics object from the bitmap

                    Using gfx As Graphics = Graphics.FromImage(bmp)
                        ' Take a screenshot of our window

                        gfx.CopyFromScreen(pt.X, pt.Y, 0, 0, sz, CopyPixelOperation.SourceCopy)

                        Dim processed As Bitmap

                        processed = bmp

                        If filename IsNot Nothing AndAlso filename <> "" Then
                            processed.Save(filename, ImageFormat.Png)

                            Active.WriteMessage("Image captured and saved to ""{0}"".", filename)
                        End If
                        processed.Dispose()
                    End Using
                End Using
                If ad.WhiteBackground Then
                    'If vtrId <> ObjectId.Null OrElse sbId <> ObjectId.Null Then
                    '    Remove3DBackground(db, tr, vtrId, sbId)
                    'Else
                    '    If gotSettings Then
                    acedSetCurrentColors(ocs)
                    Active.WriteMessage(vbLf)
                    Active.Editor.Regen()
                    'End If
                    'End If
                    Active.Editor.UpdateScreen()
                End If
                tr.Commit()
            End Using
            Return filename
        End Function

#End Region
#Region "Using Helper Functions"
        ''' <summary>
        ''' Using Transaction helper function
        ''' </summary>
        ''' <param name="action"></param>
        ''' <remarks></remarks>
        Public Sub UsingTransaction(action As Action(Of Transaction))
            Using tr = Active.Database.TransactionManager.StartTransaction()
                ' Invoke the method
                action(tr)

                tr.Commit()
            End Using
        End Sub

        ''' <summary>
        ''' Using ModelSpace helper function
        ''' </summary>
        ''' <param name="action"></param>
        ''' <remarks></remarks>
        Public Sub UsingModelSpace(action As Action(Of Transaction, BlockTableRecord))
            Using tr = Active.Database.TransactionManager.StartTransaction()
                ' Get the block table for the current database
                Dim blockTable = DirectCast(tr.GetObject(Active.Database.BlockTableId, OpenMode.ForRead), BlockTable)

                ' Get the model space block table record
                Dim modelSpace = DirectCast(tr.GetObject(blockTable(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                ' Invoke the method
                action(tr, modelSpace)

                tr.Commit()
            End Using
        End Sub

        ''' <summary>
        ''' Using PaperSpace helper function
        ''' </summary>
        ''' <param name="action"></param>
        ''' <remarks></remarks>
        Public Sub UsingPaperSpace(action As Action(Of Transaction, BlockTableRecord))
            Using tr = Active.Database.TransactionManager.StartTransaction()
                ' Get the block table for the current database
                Dim blockTable = DirectCast(tr.GetObject(Active.Database.BlockTableId, OpenMode.ForRead), BlockTable)

                ' Get the model space block table record
                Dim paperSpace = DirectCast(tr.GetObject(blockTable(BlockTableRecord.PaperSpace), OpenMode.ForRead), BlockTableRecord)

                ' Invoke the method
                action(tr, paperSpace)

                tr.Commit()
            End Using
        End Sub
#End Region

        

        

    End Class

    Public Class AppData
        Inherits ApplicationSettingsBase
        <UserScopedSetting> _
        <DefaultSettingValue("false")> _
        Public Property WhiteBackground() As Boolean
            Get
                Return CBool(Me("WhiteBackground"))
            End Get
            Set(value As Boolean)
                Me("WhiteBackground") = CBool(value)
            End Set
        End Property
        <UserScopedSetting> _
        <DefaultSettingValue("1.0")> _
        Public Property FileCaptureDelay() As Double
            Get
                Return CDbl(Me("FileCaptureDelay"))
            End Get
            Set(value As Double)
                Me("FileCaptureDelay") = CDbl(value)
            End Set
        End Property

    End Class
    ''' <summary>
    ''' Defines each Revision.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Revision
        Private rev_ChangeNote As AttRef
        ''' <summary>
        ''' The specific change note for this revision.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RevChangeNote() As AttRef
            Get
                Return rev_ChangeNote
            End Get
            Set(ByVal value As AttRef)
                rev_ChangeNote = value
            End Set
        End Property
        Private rev_Issue As AttRef
        ''' <summary>
        ''' The specific issue for this revision.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RevIssue() As AttRef
            Get
                Return rev_Issue
            End Get
            Set(ByVal value As AttRef)
                rev_Issue = value
            End Set
        End Property
        Private rev_Date As AttRef
        ''' <summary>
        ''' The specific date for this revision.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RevDate() As AttRef
            Get
                Return rev_Date
            End Get
            Set(ByVal value As AttRef)
                rev_Date = value
            End Set
        End Property
    End Class

    ''' <summary>
    ''' Our attref Object
    ''' </summary>
    ''' <remarks></remarks>
    Public Class AttRef
        Private m_attrefTag As String
        ''' <summary>
        ''' The Unique Tag Name
        ''' </summary>
        ''' <value>Driven by the AttributeReference driven into it.</value>
        ''' <returns>The Unique Tag Name</returns>
        ''' <remarks></remarks>
        Public Property attRefTag() As String
            Get
                Return m_attrefTag
            End Get
            Set(ByVal value As String)
                m_attrefTag = value
            End Set
        End Property

        Private m_textString As String
        ''' <summary>
        ''' The Unique Text String contained with the AttributeReference driven into it.
        ''' </summary>
        ''' <value>Driven by the AttributeReference driven into it.</value>
        ''' <returns>The Text String</returns>
        ''' <remarks></remarks>
        Public Property attRefText() As String
            Get
                Return m_textString
            End Get
            Set(ByVal value As String)
                m_textString = value
            End Set
        End Property
        Private m_attrefId As ObjectId
        ''' <summary>
        ''' The Unique Identifier for this AttributeReference.
        ''' </summary>
        ''' <value>Driven by the AttributeReference driven into it.</value>
        ''' <returns>The unique ObjectId we placed into it.</returns>
        ''' <remarks></remarks>
        Public Property attrefId() As ObjectId
            Get
                Return m_attrefId
            End Get
            Set(ByVal value As ObjectId)
                m_attrefId = value
            End Set
        End Property

    End Class

    ''' <summary>
    ''' Our FloatingText Object
    ''' </summary>
    ''' <remarks></remarks>
    Public Class FloatingText
        Private StrValue As String
        ''' <summary>
        ''' The String Value associated with this piece of text we found
        ''' </summary>
        ''' <value>the value to return</value>
        ''' <returns>the string object found in the area we searched.</returns>
        ''' <remarks></remarks>
        Public Property StringValue() As String
            Get
                Return StrValue
            End Get
            Set(ByVal value As String)
                StrValue = value
            End Set
        End Property
        Private StrLoc As Point3d
        ''' <summary>
        ''' The Point3d location of the piece of text we found
        ''' </summary>
        ''' <value></value>
        ''' <returns>the point3d object for this piece of FloatingText</returns>
        ''' <remarks></remarks>
        Public Property StringLocation() As Point3d
            Get
                Return StrLoc
            End Get
            Set(ByVal value As Point3d)
                StrLoc = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the height of the text object
        ''' </summary>
        ''' <remarks></remarks>
        Private txtHeight As Double
        Public Property TextHeight() As Double
            Get
                Return txtHeight
            End Get
            Set(ByVal value As Double)
                txtHeight = value
            End Set
        End Property


    End Class

    ''' <summary>
    ''' Our ZoomCommands Class
    ''' </summary>
    Public NotInheritable Class ZoomCommands

        ''' <summary>
        ''' zoom the current view using the minPoint and maxPoint
        ''' </summary>
        ''' <param name="minPoint">Point3d window origin</param>
        ''' <param name="maxPoint">Point3d window extents</param>
        ''' <remarks></remarks>
        Public Shared Sub ZoomToWindow( _
                  ByVal minPoint As Point3d, _
                  ByVal maxPoint As Point3d)
            'get the current view
            Dim view As ViewTableRecord = Active.editor.GetCurrentView()
            'start transaction
            Using trans As Transaction = Active.Database. _
              TransactionManager.StartTransaction()
                'get the entity' extends
                'configure the new current view
                view.Width = maxPoint.X - minPoint.X
                view.Height = maxPoint.Y - minPoint.Y
                view.CenterPoint = New Point2d( _
                  minPoint.X + (view.Width / 2), _
                  minPoint.Y + (view.Height / 2))
                'update the view
                Active.Editor.SetCurrentView(view)
                trans.Commit()
            End Using
        End Sub
    End Class

    ''' <summary>
    ''' Provides easy access to several "active" objects in the AutoCAD
    ''' runtime environment.
    ''' </summary>
    Public NotInheritable Class Active
        Private Sub New()
        End Sub
        ''' <summary>
        ''' Returns the active Editor object.
        ''' </summary>
        Public Shared ReadOnly Property Editor() As Editor
            Get
                Return Document.Editor
            End Get
        End Property

        ''' <summary>
        ''' Returns the active Document object.
        ''' </summary>
        Public Shared ReadOnly Property Document() As Document
            Get
                Return Application.DocumentManager.MdiActiveDocument
            End Get
        End Property

        ''' <summary>
        ''' Returns the active Database object.
        ''' </summary>
        Public Shared ReadOnly Property Database() As Database
            Get
                Return Document.Database
            End Get
        End Property

        ''' <summary>
        ''' Sends a string to the command line in the active Editor
        ''' </summary>
        ''' <param name="message">The message to send.</param>
        Public Shared Sub WriteMessage(message As String)
            Editor.WriteMessage(message)
        End Sub

        ''' <summary>
        ''' Sends a string to the command line in the active Editor using String.Format.
        ''' </summary>
        ''' <param name="message">The message containing format specifications.</param>
        ''' <param name="parameter">The variables to substitute into the format string.</param>
        Public Shared Sub WriteMessage(message As String, ParamArray parameter As Object())
            Editor.WriteMessage(message, parameter)
        End Sub
    End Class


End Namespace