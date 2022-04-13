Attribute VB_Name = "GarmentSim"
Option Private Module
Public Cncl As Boolean
Public GarmentPath As String
Public GarmentPathCZ As String

Sub GSLaunch()
GarmentPath = GarmentSeek
GarmentPathCZ = "S:\XI-Online\!LIVE-ACCESS\CZ-Customization"
frmGarm.Show vbModeless
End Sub
Sub GarmentSim()
Dim artDoc As Document, tempDoc As Document, Template(5) As Document, allDocs As Boolean
Dim s As Shape, xPath As String, GarmSKU As String, x As Double, y As Double
Dim dCount As Integer, opCount As Integer, tCount As Integer, d As Integer
Set fso = New FileSystemObject

If Not Right(GarmentPath, 1) = "\" Then GarmentPath = GarmentPath & "\"
If frmGarm.cbAllDocs = True Then allDocs = True
Set artDoc = ActiveDocument

If allDocs = True Then
    dCount = Documents.Count                   'dcount should be documents count - templates(docs with property) count??
    OpenCMYK = cmyBool: OpenWCMYK = wcmBool    'might as well get total ops in get dcount function for total percentage on bar??
Else
    dCount = 1
    If Left(ActiveDocument.Name, 1) = "W" Then
        OpenWCMYK = True
    Else
        OpenCMYK = True
    End If
End If

If OpenCMYK = True Then
    Set Template(0) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\CMYK-1.cdr")
    Set Template(1) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\CMYK-2.cdr")
    Set Template(2) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\CMYK-3.cdr")
End If
If OpenWCMYK = True Then
    Set Template(3) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\WCMYK-1.cdr")
    Set Template(4) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\WCMYK-2.cdr")
    Set Template(5) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\WCMYK-3.cdr")
End If

For d = 1 To dCount
If allDocs Then Set artDoc = Documents(d)
    artDoc.DrawingOriginX = 0: Ox = artDoc.pages(1).SizeWidth
    artDoc.DrawingOriginY = 0: Oy = artDoc.pages(1).SizeHeight
    artDoc.pages(1).SelectShapesFromRectangle(-Ox / 2, Oy / 2, Ox / 2, -Oy / 2, True).Group.Copy
    artDoc.Dirty = False
    
    GarmSKU = fso.GetBaseName(artDoc.FullFileName)
    If Left(GarmSKU, 1) = "W" Then
        tCount = 3: opCount = 0
    Else
        tCount = 0: opCount = 0
    End If
    
    For t = tCount To tCount + 2
        Set tempDoc = Template(t)
        ShapeForGarm tempDoc.MasterPage.Layers("Design").Paste, tempDoc
        tempDoc.Activate
        
        For p = 1 To tempDoc.pages.Count
            opCount = opCount + 1
            tempDoc.pages(p).Activate ': DoEvents
            xPath = GarmentPath & GarmSKU & "-" & tempDoc.pages(p).Name & ".jpg"
            
            tempDoc.pages(p).Layers("Layer 1").Shapes.All.GetSize x, y
            FindDims x, y
            
                tempDoc.ExportBitmap(xPath, cdrJPEG, cdrCurrentPage, cdrRGBColorImage, x, y).Finish
                UpdateBar GarmSKU, tCount, opCount, d - 1, dCount: DoEvents
                If Cncl Then GarmCancel frmGarm.BarCZ, frmGarm.GarmCancelCZ: Exit Sub
                'If opCount Mod 6 = 0 Then DoEvents
        Next p
    Next t
    
    UpdateBar GarmSKU, tCount, opCount, d, dCount
Next d

For i = 0 To UBound(Template())
    If Not Template(i) Is Nothing Then Template(i).Close
Next i

Set fso = Nothing
End Sub
Sub UpdateBar(GarmSKU As String, tCount As Integer, opCount As Integer, d As Integer, dCount As Integer)
If tCount = 0 Then
    total = 139
ElseIf tCount = 3 Then
    total = 156
End If

frmGarm.Bar.Width = opCount * (253 / total)
frmGarm.ProgCounter.Caption = GarmSKU & " (" & opCount & "/" & total & ")"
frmGarm.ProgTotal.Caption = "Total Docs (" & d & "/" & dCount & ")"

frmGarm.ProgTotal.Visible = True
frmGarm.ProgCounter.Visible = True
'frmGarm.Repaint
'DoEvents
End Sub
Sub GarmentSimCZ()
Dim artDoc As Document, tempDoc As Document, Template(5) As Document, allDocs As Boolean
Dim s As Shape, sr As ShapeRange, xPath As String, GarmSKU As String, Ox As Double, Oy As Double
Dim dCount As Integer, opCount As Integer, tCount As Integer, czCount As Integer, d As Integer

Dim czsample As Boolean, cztext As Boolean, czcoords As Boolean, czblank As Boolean
Dim czlayer As Layer, simLayer As Layer, tmpLayer As Layer, br As New ShapeRange
Set fso = New FileSystemObject

If Not Right(GarmentPath, 1) = "\" Then GarmentPathCZ = GarmentPathCZ & "\"
If frmGarm.cbAllDocsCZ = True Then allDocs = True
Set artDoc = ActiveDocument

czCount = GetCZCount(czsample, cztext, czcoords, czblank)

If allDocs = True Then
    dCount = Documents.Count                   'dcount should be documents count - templates(docs with property) count??
    OpenCMYK = cmyBool: OpenWCMYK = wcmBool    'might as well get total ops in get dcount function for total percentage on bar??
Else
    dCount = 1
    If Left(ActiveDocument.Name, 1) = "W" Then
        OpenWCMYK = True
    Else
        OpenCMYK = True
    End If
End If

If OpenCMYK = True Then
    Set Template(0) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\CMYK-1.cdr")
    Set Template(1) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\CMYK-2.cdr")
    Set Template(2) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\CMYK-3.cdr")
End If
If OpenWCMYK = True Then
    Set Template(3) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\WCMYK-1.cdr")
    Set Template(4) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\WCMYK-2.cdr")
    Set Template(5) = OpenDocument("S:\XI-Online\!-All Other Stuff\Graphic Artist References\TEMPLATES\All Adult Garments\WCMYK-3.cdr")
End If

For d = 1 To dCount
If allDocs Then Set artDoc = Documents(d)
    For ly = 1 To artDoc.pages(1).Layers.Count
        If artDoc.pages(1).Layers(ly).Name = "Custom Text" Then Set czlayer = artDoc.pages(1).Layers(ly): Exit For
    Next ly
    
    If Not czlayer Is Nothing Then
        Set sr = czlayer.FindShapes(Type:=cdrTextShape, recursive:=False)
        
        If sr.Count = 0 Then
            MsgBox "No Custom Text objects found in " & artDoc.Title
            GoTo Skip
        End If
    Else
        MsgBox "No Custom Text layer found in " & artDoc.Title
        GoTo Skip
    End If
    
    artDoc.DrawingOriginX = 0: Ox = artDoc.pages(1).SizeWidth
    artDoc.DrawingOriginY = 0: Oy = artDoc.pages(1).SizeHeight
    
    GarmSKU = fso.GetBaseName(artDoc.FullFileName)
    If Left(GarmSKU, 1) = "W" Then
        tCount = 3: opCount = 0
    Else
        tCount = 0: opCount = 0
    End If
    
    Set tmpLayer = artDoc.pages(1).CreateLayer("Temp")
    sr.CopyToLayer tmpLayer
    tmpLayer.Shapes.All.Move Ox, Oy
    
    Set simLayer = artDoc.pages(1).CreateLayer("Custom Sim")
    If czsample Then
        'Sample - Export pics like normal
        SimPrepCZ artDoc, simLayer, Ox, Oy
        
        For t = tCount To tCount + 2
            ExportGarmCZ Template(t), GarmSKU, tCount, opCount, d, dCount, czCount, "-sample", czcoords, Template(t).pages.Count
            If Cncl Then Exit Sub
        Next t
    End If
    If cztext Then
        'Text - Change text.story to "text"
        For Each s In sr
            s.Text.Story = "TEXT"
        Next s
        
        SimPrepCZ artDoc, simLayer, Ox, Oy
        
        For t = tCount To tCount + 2
            ExportGarmCZ Template(t), GarmSKU, tCount, opCount, d, dCount, czCount, "-text", czcoords, Template(t).pages.Count
            If Cncl Then Exit Sub
        Next t
    End If
    If czblank Then
        'Blank - Make text invisible
        For Each s In sr
            s.Fill.ApplyNoFill
            s.Outline.SetNoOutline
        Next s
            
        SimPrepCZ artDoc, simLayer, Ox, Oy
        
        For t = tCount To tCount + 2
            ExportGarmCZ Template(t), GarmSKU, tCount, opCount, d, dCount, czCount, "-blank", czcoords, Template(t).pages.Count
            If Cncl Then Exit Sub
        Next t
    End If
    If czcoords Then
        'Coord - Create bound boxes
        For Each s In sr
            br.Add CreatePSBound(s, artDoc, czlayer)
        Next s
        
        SimPrepCZ artDoc, simLayer, Ox, Oy
        br.Delete
        
        'no loop - only need 1 coord
        ExportGarmCZ Template(tCount), GarmSKU, tCount, opCount, d, dCount, czCount, "-COORDS", czcoords
        If Cncl Then Exit Sub
    End If
    
    simLayer.Delete
    sr.Delete
    
    tmpLayer.Shapes.All.Move -Ox, -Oy
    tmpLayer.Shapes.All.CopyToLayer czlayer
    tmpLayer.Delete
    artDoc.Dirty = False
    
    UpdateBarCZ GarmSKU, tCount, opCount, d, dCount, czCount, czcoords
    
Skip:
Next d

For i = 0 To UBound(Template())
    If Not Template(i) Is Nothing Then Template(i).Close
Next i

Set fso = Nothing
End Sub
Sub ExportGarmCZ(tempDoc As Document, GarmSKU As String, tCount As Integer, ByRef opCount As Integer, d As Integer, dCount As Integer, czCount As Integer, suff As String, czcoords As Boolean, Optional pages As Integer = 1)
Dim x As Double, y As Double, xFil As cdrFilter, xTn As String, bool400 As Boolean, pref As String
pref = ""

    ShapeForGarm tempDoc.MasterPage.Layers("Design").Paste, tempDoc
    tempDoc.Activate
    
    If suff = "-blank" Or suff = "-COORDS" Then
        bool400 = True
        x = 400
        y = 400
        xFil = cdrPNG
        xTn = ".png"
        If suff = "-COORDS" Then
            pref = "!"
        End If
    Else
        xFil = cdrJPEG
        xTn = ".jpg"
    End If
    
    For p = 1 To pages
        If bool400 = True Then
            If Right(tempDoc.Properties("Type", 1), 5) <> "Sim 3" Then
                tempDoc.pages(p).Layers("400x400").Printable = True
            End If
        Else
            tempDoc.pages(p).Layers("Layer 1").Shapes(1).GetSize x, y
            FindDims x, y
        End If
        
        opCount = opCount + 1
        tempDoc.pages(p).Activate ': DoEvents
        xPath = GarmentPathCZ & pref & GarmSKU & suff & "-" & tempDoc.pages(p).Name & xTn
        
            tempDoc.ExportBitmap(xPath, xFil, cdrCurrentPage, cdrRGBColorImage, x, y).Finish
            If bool400 And Right(tempDoc.Properties("Type", 1), 5) <> "Sim 3" Then tempDoc.pages(p).Layers("400x400").Printable = False
        
        UpdateBarCZ GarmSKU, tCount, opCount, d - 1, dCount, czCount, czcoords: DoEvents
        If Cncl Then GarmCancel frmGarm.Bar, frmGarm.GarmCancel: Exit Sub
        'If opCount Mod 6 = 0 Then DoEvents
    Next p
End Sub
Sub UpdateBarCZ(GarmSKU As String, tCount As Integer, opCount As Integer, d As Integer, dCount As Integer, czCount As Integer, czcoords As Boolean)
If tCount = 0 Then
    total = 139 * czCount
ElseIf tCount = 3 Then
    total = 156 * czCount
End If
If czcoords Then total = total + 1

frmGarm.BarCZ.Width = opCount * (253 / total)
frmGarm.ProgCounterCZ.Caption = GarmSKU & " (" & opCount & "/" & total & ")"
frmGarm.ProgTotalCZ.Caption = "Total Docs (" & d & "/" & dCount & ")"

frmGarm.ProgTotalCZ.Visible = True
frmGarm.ProgCounterCZ.Visible = True
'frmGarm.Repaint
'DoEvents
End Sub
Sub AddSelCZ()
Dim s As Shape, sr As ShapeRange, czlayer As Layer

Set sr = ActiveSelectionRange.Shapes.FindShapes(Type:=cdrTextShape, recursive:=False)
If sr.Count = 0 Then
    MsgBox "No text in selection": Exit Sub
End If

For i = 1 To ActivePage.Layers.Count
    If ActivePage.Layers(i).Name = "Custom Text" Then Set czlayer = ActivePage.Layers(i): Exit For
Next i

ActiveDocument.BeginCommandGroup
    If czlayer Is Nothing Then
        'MsgBox "Custom Text layer not found. Creating."
        Set czlayer = ActivePage.CreateLayer("Custom Text")
    End If
    
    For i = 1 To sr.Count
        sr.Shapes(i).MoveToLayer czlayer
    Next i
    
    ShowCZText
ActiveDocument.EndCommandGroup
End Sub
Sub ShowCZText()
Dim s As Shape, sr As ShapeRange, czlayer As Layer

With frmGarm
    .CZBox.Clear

    For i = 1 To ActivePage.Layers.Count
        If ActivePage.Layers(i).Name = "Custom Text" Then Set czlayer = ActivePage.Layers(i): Exit For
    Next i
    
    If Not czlayer Is Nothing Then
        Set sr = czlayer.FindShapes(Type:=cdrTextShape, recursive:=False)
        
        If sr.Count > 0 Then
            For Each s In sr
                .CZBox.AddItem s.Text.Story
            Next s
        End If
    End If
End With
End Sub
Sub SimPrepCZ(artDoc As Document, simLayer As Layer, x As Double, y As Double)
    simLayer.Activate
    artDoc.pages(1).SelectShapesFromRectangle(-x / 2, y / 2 + 0.25, x / 2, -y / 2 - 0.25, False).CopyToLayerAsRange simLayer
    artDoc.ClearSelection
    simLayer.Shapes.All.Group.Cut
End Sub
Function CreatePSBound(cs As Shape, cdoc As Document, cl As Layer) As Shape
Dim s As Shape, nr As New ShapeRange, ogTxt As String, txtBool As Boolean
Dim x#, y#, w#, h#

Dim cy As New Color, mg As New Color
cy.RGBAssign 0, 255, 255
mg.RGBAssign 255, 0, 255

If cs.Type = cdrTextShape Then
    ogTxt = cs.Text.Story
    cs.Text.Story = "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
    txtBool = True
End If

cs.GetBoundingBox x, y, w, h
Set s = cl.CreateRectangle2(x, y, w, h)
    s.Fill.UniformColor.CopyAssign mg
    s.Outline.SetNoOutline
    nr.Add s
w = w / 2
h = h / 2
Set s = cl.CreateRectangle2(x, y, w, h)
    s.Fill.UniformColor.CopyAssign cy
    s.Outline.SetNoOutline
    nr.Add s
Set s = cl.CreateRectangle2(x + w, y + h, w, h)
    s.Fill.UniformColor.CopyAssign cy
    s.Outline.SetNoOutline
    nr.Add s
    
'ActiveDocument.ClearSelection
'cl.Activate
nr.AddToSelection
Set CreatePSBound = nr.Group
    CreatePSBound.Name = "PS Coords Box"
    's.AddToSelection

If txtBool = True Then
    cs.Text.Story = ogTxt
End If

End Function
Function GetCZCount(ByRef czsample As Boolean, ByRef cztext As Boolean, ByRef czcoords As Boolean, ByRef czblank As Boolean) As Integer
GetCZCount = 0
If frmGarm.cbSample = True Then
    czsample = True
    GetCZCount = GetCZCount + 1
End If
If frmGarm.cbText = True Then
    cztext = True
    GetCZCount = GetCZCount + 1
End If
If frmGarm.cbCoords = True Then
    czcoords = True
    'GetCZCount = GetCZCount + 1
End If
If frmGarm.cbBlank = True Then
    czblank = True
    GetCZCount = GetCZCount + 1
End If
End Function
Sub GarmCancel(Bar As Object, cancelbtn As Object)
Bar.Picture = LoadPicture("S:\XI-Online\!-All Other Stuff\Ethan References\Macros\GarmentSim\BarCancel.jpg")
cancelbtn.Enabled = False
End Sub
Sub FindDims(ByRef x As Double, ByRef y As Double, Optional minDim As Integer = 1800)
Dim c As Double

If x > y Then
    c = x / y
    y = minDim
    x = minDim * c
Else
    c = y / x
    x = minDim
    y = minDim * c
End If
End Sub
Sub ShapeForGarm(s As Shape, doc As Document)
If doc.Properties.Exists("Scale", 1) Then
    If doc.MasterPage.Layers("Design").Shapes.Count > 1 Then doc.MasterPage.Layers("Design").Shapes(2).Delete
    s.SizeHeight = s.SizeHeight * doc.Properties("Scale", 1)
    s.SizeWidth = s.SizeWidth * doc.Properties("Scale", 1)
    s.AlignToShape cdrAlignHCenter, doc.MasterPage.Layers("Guides").Shapes(2)
    s.AlignToShape cdrAlignTop, doc.MasterPage.Layers("Guides").Shapes(1)
Else
    If doc.MasterPage.Layers("Design").Shapes.Count > 2 Then doc.MasterPage.Layers("Design").Shapes(2).Delete
    If s.SizeWidth > s.SizeHeight Then
        c = 3 / s.SizeWidth
        s.SizeHeight = s.SizeHeight * c
        s.SizeWidth = 3
    Else
        c = 3 / s.SizeHeight
        s.SizeWidth = s.SizeWidth * c
        s.SizeHeight = 3
    End If
        s.AlignToShape cdrAlignHCenter + cdrAlignVCenter, doc.MasterPage.Layers("Design").Shapes(2) 's.AlignToPageCenter
End If
End Sub
Function cmyBool() As Boolean
cmyBool = False

For i = 1 To Documents.Count
    If Not Left(Documents(i).Name, 1) = "W" Then
        cmyBool = True: Exit Function
    End If
Next i
End Function
Function wcmBool() As Boolean
wcmBool = False

For i = 1 To Documents.Count
    If Left(Documents(i).Name, 1) = "W" Then
        wcmBool = True: Exit Function
    End If
Next i
End Function
Function GarmentSeek() As String
Dim CrFldr As Object
Set fso = New FileSystemObject

BaseName = fso.GetBaseName(ActiveDocument)
If Left(BaseName, 1) = "W" Then BaseName = Right(BaseName, Len(BaseName) - 1)

If Left(BaseName, 2) = "CZ" Then
    GarmentSeek = "Next tab for custom"
    Set fso = Nothing: Exit Function
End If

Search1 = "S:\XI-Online\!LIVE-ACCESS\": Search2 = Left(BaseName, 2) & "*?*"

GarmentSeek = CorelScriptTools.FindFirstFolder(Search1 & Search2, 16 Or 128)
If GarmentSeek <> "" Then
    Set CrFldr = fso.GetFolder(Search1 & GarmentSeek)
    GarmentSeek = Search1 & fso.GetFileName(CrFldr)
    Set fso = Nothing: Exit Function
End If

Set fso = Nothing
GarmentSeek = "Live Access folder not found"
End Function
Function GarmChooseFolder(Optional browsePath As String = "S:\XI-Online") As String
Set fso = New FileSystemObject
  Dim shell, folder
  Set shell = CreateObject("Shell.Application")
  Set folder = shell.BrowseForFolder(0, "Select Location", &H4000 + &H10, browsePath & "\")

    If Not folder Is Nothing Then
          GarmChooseFolder = folder.self.path & "\"
    Else: GarmChooseFolder = ""
    End If

  Set folder = Nothing
  Set shell = Nothing
  Set fso = Nothing
End Function
