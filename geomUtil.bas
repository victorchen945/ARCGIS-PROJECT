Attribute VB_Name = "geomutil"
' Geometry Utility Library
'

Option Explicit
'------------------------------------------------------------------------------------
Public Function Cylinder( _
    pOrigin As IPoint, _
    radius As Double, _
    minLon As Double, _
    maxLon As Double, _
    zmin As Double, _
    zmax As Double, _
    Optional bFlipS As Boolean = False, _
    Optional bFlipT As Boolean = False) As IMultiPatch
    
    On Error GoTo EH
    
    Dim sampLon As Double
     sampLon = 36
    
    Dim xStep As Double
    Dim yStep As Double
    xStep = (maxLon - minLon) / sampLon
    
    Dim lonRange As Double
     lonRange = maxLon - minLon
    
    Dim pMultiPatch As IMultiPatch
     Set pMultiPatch = New MultiPatch
    
    Dim pGCol As IGeometryCollection
     Set pGCol = pMultiPatch
    
    Dim pGeom As IGeometry2
    
    Dim pt As esricore.IPoint
    
    Dim pStrip As IPointCollection
     Set pStrip = New TriangleStrip
    
    Dim pVector As IVector3D
     Set pVector = New Vector3D
    
    Dim pGE As IEncode3DProperties
     Set pGE = New GeometryEnvironment
    
    Dim lon As Double
     For lon = maxLon To minLon Step -xStep
        Dim azi As Double
         azi = DegreesToRadians(lon)
         pVector.PolarSet -azi, 0, radius
         Set pt = New esricore.Point
         pt.x = pOrigin.x + pVector.XComponent
         pt.y = pOrigin.y + pVector.YComponent
         pt.Z = zmin
        Dim s As Double
         s = (lon - minLon) / lonRange
         If (bFlipS) Then
             s = 1 + (s * -1)
         End If
        ' Due to floating point precision issues make sure
        ' texture coordinate in safe range.
         If (s <= 0) Then
             s = 0.001
             ElseIf (s >= 1) Then
             s = 0.999
         End If
        
        Dim t As Double
         If (bFlipT) Then
             t = 0
         Else
             t = 1
         End If
        
        ' pack the s/t into same measure used for vector normal
        ' and assign the measure to the point
        Dim m As Double
         m = 0
         pGE.PackTexture2D s, t, m
         pt.m = m
         pStrip.AddPoint pt
        
        Dim pt2 As IPoint
        Dim pClone As IClone
         Set pClone = pt
         Set pt2 = pClone.Clone
         pt2.Z = zmax
         If (bFlipT) Then
             t = 1
         Else
             t = 0
         End If
         m = 0
         pGE.PackTexture2D s, t, m
         pt2.m = m
         pStrip.AddPoint pt2
     Next lon
    
     Set pGeom = pStrip
     pGCol.AddGeometry pGeom
    
    Dim pZAware As IZAware
     Set pZAware = pMultiPatch
     pZAware.ZAware = True
    
    Dim pMAware As IMAware
     Set pMAware = pMultiPatch
     pMAware.MAware = True
    
     Set Cylinder = pMultiPatch
    Exit Function
EH:
     Set Cylinder = Nothing
End Function

'------------------------------------------------------------------------------------
Public Function Sphere( _
    minLon As Double, _
    maxLon As Double, _
    minLat As Double, _
    maxLat As Double, _
    origin As esricore.IPoint, _
    radius As Double, _
    Optional bSmooth As Boolean = False, _
    Optional bFlipS As Boolean = False, _
    Optional bFlipT As Boolean = False, _
    Optional lDivision As Long = 36) As IMultiPatch
    
    On Error GoTo EH
    
    Dim sampLon As Double
     sampLon = lDivision
    
    Dim xStep As Double
    Dim yStep As Double
     xStep = (maxLon - minLon) / sampLon
     yStep = (maxLat - minLat) / (sampLon / 2)
    
    Dim lonRange As Double
     lonRange = maxLon - minLon
    
    Dim latRange As Double
     latRange = maxLat - minLat
    
    Dim pMultiPatch As IMultiPatch
     Set pMultiPatch = New MultiPatch
    
    Dim pGCol As IGeometryCollection
     Set pGCol = pMultiPatch
    
    Dim pGeom As IGeometry2
    
    Dim pt As esricore.IPoint
    
    Dim pStrip As IPointCollection
    
    Dim pVector As IVector3D
     Set pVector = New Vector3D
    
    Dim pGE As IEncode3DProperties
     Set pGE = New GeometryEnvironment
    
    Dim lon As Double
     For lon = minLon To (maxLon - xStep) Step xStep
         Set pStrip = New TriangleStrip
        Dim lat As Double
         For lat = minLat To maxLat Step yStep
            Dim azi As Double
            Dim inc As Double
             azi = DegreesToRadians(lon)
             inc = DegreesToRadians(lat)
             pVector.PolarSet -azi, inc, radius
             Set pt = New esricore.Point
             pt.x = origin.x + pVector.XComponent
             pt.y = origin.y + pVector.YComponent
             pt.Z = origin.Z + pVector.ZComponent
            Dim s As Double
             s = (lon - minLon) / lonRange
             If (bFlipS) Then s = 1 + (s * -1)
            ' Due to floating point precision issues make sure
            ' texture coordinate in safe range.
             If (s <= 0) Then
                 s = 0.001
                 ElseIf (s >= 1) Then
                 s = 0.999
             End If
            
            Dim t As Double
             t = (maxLat - lat) / latRange
             If (bFlipT) Then t = 1 + (t * -1)
             If (t <= 0) Then
                 t = 0.001
                 ElseIf (t >= 1) Then
                 t = 0.999
             End If
            
            ' pack the s/t into same measure used for vector normal
            ' and assign the measure to the point
            Dim m As Double
             m = 0
             pGE.PackTexture2D s, t, m
             If (bSmooth) Then     'pack normal for smoothing
                 pVector.Normalize
                 pGE.PackNormal pVector, m
             End If
             pt.m = m
             pStrip.AddPoint pt
             If ((lat <> -90) And (lat <> 90)) Then
                 azi = DegreesToRadians(lon + xStep)
                 inc = DegreesToRadians(lat)
                 pVector.PolarSet -azi, inc, radius
                 Set pt = New esricore.Point
                 pt.x = origin.x + pVector.XComponent
                 pt.y = origin.y + pVector.YComponent
                 pt.Z = origin.Z + pVector.ZComponent
                 s = (lon + xStep - minLon) / lonRange
                 If (bFlipS) Then s = 1 + (s * -1)
                 If (s <= 0) Then
                     s = 0.001
                     ElseIf (s >= 1) Then
                     s = 0.999
                 End If
                
                 t = (maxLat - lat) / latRange
                 If (bFlipT) Then t = 1 + (t * -1)
                 If (t <= 0) Then
                     t = 0.001
                     ElseIf (t >= 1) Then
                     t = 0.999
                 End If
                
                 m = 0
                 pGE.PackTexture2D s, t, m
                 If (bSmooth) Then       'pack normal for smoothing
                     pVector.Normalize
                     pGE.PackNormal pVector, m
                 End If
                 pt.m = m
                 pStrip.AddPoint pt
             End If
         Next lat
         Set pGeom = pStrip
         pGCol.AddGeometry pGeom
     Next lon
    
    Dim pMAware As IMAware
     Set pMAware = pMultiPatch
     pMAware.MAware = True
    
     Set Sphere = pMultiPatch
    Exit Function
EH:
     Set Sphere = Nothing
End Function

'------------------------------------------------------------------------------------
Public Function DegreesToRadians(dDeg As Double) As Double
    Dim PI As Double
     PI = 4 * Atn(1#)
    
    Dim RAD As Double
     RAD = PI / 180#
    
     DegreesToRadians = dDeg * RAD
End Function


'Ranging from a top-n-bottom bottomless cube to a cylinder, depending on divisions.
'If OffCenter is provided, the polyhedron can be slant/tilted.
'------------------------------------------------------------------------------------
Public Function Polyhedron( _
    pOrigin As esricore.IPoint, _
    radius As Double, _
    minLon As Double, _
    maxLon As Double, _
    zmin As Double, _
    zmax As Double, _
    iDivision As Integer, _
    Optional pOffCenter As esricore.IPoint = Nothing, _
    Optional bSmooth As Boolean = False, _
    Optional bFlipS As Boolean = False, _
    Optional bFlipT As Boolean = False) As esricore.IMultiPatch
    
    On Error GoTo EH
    
    Dim sampLon As Double
     sampLon = iDivision
    
    Dim xStep As Double
    Dim yStep As Double
     xStep = (maxLon - minLon) / sampLon
    
    Dim lonRange As Double
     lonRange = maxLon - minLon
    
    Dim pMultiPatch As IMultiPatch
     Set pMultiPatch = New MultiPatch
    
    Dim pGCol As IGeometryCollection
     Set pGCol = pMultiPatch
    
    Dim pGeom As IGeometry2
    
    Dim pt As esricore.IPoint
    
    Dim pStrip As IPointCollection
     Set pStrip = New TriangleStrip
    
    Dim pVector As IVector3D
     Set pVector = New Vector3D
    
    Dim pGE As IEncode3DProperties
     Set pGE = New GeometryEnvironment
    
    Dim lon As Double
     For lon = maxLon To minLon Step -xStep
        Dim azi As Double
         azi = DegreesToRadians(lon)
         pVector.PolarSet -azi, 0, radius
         Set pt = New esricore.Point
         pt.x = pOrigin.x + pVector.XComponent
         pt.y = pOrigin.y + pVector.YComponent
        
         If (pOffCenter Is Nothing) Then     'apply possible smoothing when upright
            Dim m As Double
             m = 0
             If (bSmooth) Then
                Dim pV As esricore.IVector3D
                 Set pV = New Vector3D
                 pV.SetComponents pt.x, pt.y, 0
                 pV.Normalize
                 pGE.PackNormal pV, m
                 pt.m = m
             End If
         Else
             If (pOffCenter.Z = pOrigin.Z) Then
                 MsgBox "The two point cannot be at the same height."
                 Set Polyhedron = Nothing
                Exit Function
             End If
             If Not (pOffCenter.Z = zmin Or pOffCenter.Z = zmax) Then
                 MsgBox "The second point height has to be the same as zMin or zMax."
                 Set Polyhedron = Nothing
                Exit Function
             End If
            Dim pSubV As esricore.IVector3D
             Set pSubV = New Vector3D
             pSubV.ConstructDifference pOffCenter, pOrigin
             pt.x = pt.x + pSubV.XComponent
             pt.y = pt.y + pSubV.YComponent
         End If
         pt.Z = zmin
        Dim s As Double
         s = (lon - minLon) / lonRange
         If (bFlipS) Then s = 1 + (s * -1)
        ' Due to floating point precision issues make sure
        ' texture coordinate in safe range.
         If (s <= 0) Then s = 0.001
         If (s >= 1) Then s = 0.999
        
        Dim t As Double
         t = IIf(bFlipT, 0, 1)
         pGE.PackTexture2D s, t, m
         pt.m = m
         pStrip.AddPoint pt
        
        Dim pt2 As IPoint
        Dim pClone As IClone
         Set pClone = pt
         Set pt2 = pClone.Clone
         If (pOffCenter Is Nothing) Then     'apply possible smoothing when upright
             If (bSmooth) Then pt2.m = m     'use the above calculated m
         Else
            Dim pSubV2 As esricore.IVector3D
             Set pSubV2 = New Vector3D
             pSubV2.ConstructDifference pOrigin, pOffCenter
             pt2.x = pt2.x + pSubV2.XComponent
             pt2.y = pt2.y + pSubV2.YComponent
         End If
         pt2.Z = zmax
         t = IIf(bFlipT, 1, 0)
         m = 0
         pGE.PackTexture2D s, t, m
         pt2.m = m
         pStrip.AddPoint pt2
     Next lon
    
     Set pGeom = pStrip
     pGCol.AddGeometry pGeom
    
    Dim pZAware As IZAware
     Set pZAware = pMultiPatch
     pZAware.ZAware = True
    
    Dim pMAware As IMAware
     Set pMAware = pMultiPatch
     pMAware.MAware = True
    
     Set Polyhedron = pMultiPatch
    Exit Function
EH:
     Set Polyhedron = Nothing
     MsgBox Err.Number & ": " & Err.Description & " in GeomUtil.Polyhedron()"
End Function

'Ranging from a four-side pyramid to a cone, depending on the number of divisions.
'If pCenter is provided, then the cone could be made slant/tilted.
'No texture and normal packing.
'------------------------------------------------------------------------------------
Public Function Pyramid( _
    pTop As esricore.IPoint, _
    radius As Double, _
    minLon As Double, _
    maxLon As Double, _
    zPlaneHeight As Double, _
    iDivision As Integer, _
    Optional pCenter As esricore.IPoint = Nothing) As esricore.IMultiPatch
    
    On Error GoTo EH
    
    Dim sampLon As Double
     sampLon = iDivision
    Dim dblStep As Double
     dblStep = (maxLon - minLon) / sampLon
    
    Dim pMP As esricore.IMultiPatch
     Set pMP = New MultiPatch
    Dim pGCol As esricore.IGeometryCollection
     Set pGCol = pMP
    
    Dim pFan As esricore.IPointCollection
     Set pFan = New TriangleFan
     pFan.AddPoint pTop
    Dim pV3D As esricore.IVector3D
     Set pV3D = New Vector3D
    
    Dim dblAngle As Double                      'the inclination angle from top pt
     dblAngle = IIf(pTop.Z = 0, 0, Atn(radius / pTop.Z))
    
    Dim lon As Double
    Dim azi As Double
    Dim pt As esricore.IPoint
     For lon = maxLon To minLon Step -dblStep     'pay attention to the order
         azi = DegreesToRadians(lon)
         pV3D.PolarSet azi, dblAngle, radius
         Set pt = New Point
         pt.x = pTop.x + pV3D.XComponent
         pt.y = pTop.y + pV3D.YComponent
         If Not (pCenter Is Nothing) Then
             If (pCenter.Z <> zPlaneHeight) Then
                 MsgBox "The plane cannot be tilted."
                 Set Pyramid = Nothing
                Exit Function
             End If
             pt.x = pt.x + pCenter.x
             pt.y = pt.y + pCenter.y
         End If
         pt.Z = zPlaneHeight
         pFan.AddPoint pt
     Next lon
     pGCol.AddGeometry pFan
    
    Dim pZA As esricore.IZAware
     Set pZA = pMP
     pZA.ZAware = True
    Dim pMA As esricore.IMAware
     Set pMA = pMP
     pMA.MAware = True
    
     Set Pyramid = pMP
    
    Exit Function
EH:
     Set Pyramid = Nothing
     MsgBox Err.Number & ": " & Err.Description & " in GeomUtil.Pyramid()"
End Function

'Requires a series of points as parameters (not an array).  The order of points has _
to be
'in such a way that consecutive triangles could be formed.
'No texture and normal packing.
'------------------------------------------------------------------------------------
Public Function Prismaid(ParamArray Points() As Variant) As esricore.IMultiPatch
    On Error GoTo EH
     If Not IsArray(Points) Then
         MsgBox "Parameter array required."
         Set Prismaid = Nothing
        Exit Function
     Else
         If Not (TypeOf Points(0) Is IPoint) Then
             MsgBox "Point array required."
             Set Prismaid = Nothing
            Exit Function
         End If
        Dim pPtItem
        Dim iCount As Integer
         iCount = 0
         For Each pPtItem In Points
             iCount = iCount + 1
         Next
         If (iCount Mod 2 = 1) Then  'odd number of points
             MsgBox "Even number of points required."
             Set Prismaid = Nothing
            Exit Function
         End If
     End If
    
    Dim pStrip As esricore.IPointCollection
     Set pStrip = New esricore.TriangleStrip
    
    Dim i As Integer
     For i = 0 To iCount - 1
         pStrip.AddPoint Points(i)
     Next i
    
    Dim pGeometry As esricore.IGeometry2
     Set pGeometry = pStrip
    Dim pMultiPatch As esricore.IMultiPatch
     Set pMultiPatch = New MultiPatch
    Dim pGeoCol As esricore.IGeometryCollection
     Set pGeoCol = pMultiPatch
     pGeoCol.AddGeometry pGeometry
    
    Dim pZAware As esricore.IZAware
     Set pZAware = pMultiPatch
     pZAware.ZAware = True
    Dim pMAware As esricore.IMAware
     Set pMAware = pMultiPatch
     pMAware.MAware = True
    
     Set Prismaid = pMultiPatch
    
    Exit Function
EH:
     MsgBox Err.Number & ": " & Err.Description & " in GeomUtil.Prismaid()"
End Function

' AlongPolyLineZ - Returns a list of 3D points that are spaced
' along a 3D polyline at a given interval. The interval is in
' true 3D distance. An optional vector list may be returned that
' provides segment direction at sample positions.
'
' Current design limit: only processes first part/path of input polyline.
'------------------------------------------------------------------------------------
Public Function AlongPolyLineZ( _
    ByVal ln As IPolyline, _
    ByVal step As Double, _
    ByVal doVectors As Boolean, _
    ByRef pVectors As IArray) As IPointCollection
    
    Dim pGC As IGeometryCollection
     Set pGC = ln
    
    Dim pInPC As IPointCollection
     Set pInPC = pGC.Geometry(0)
    
    Dim pt1 As IPoint
     Set pt1 = pInPC.Point(0)
    
    Dim pOutPC As IPointCollection
     Set pOutPC = New Path
    
    ' Make all the points added to the collection Z aware
    Dim pZAware As IZAware
     Set pZAware = pOutPC
     pZAware.ZAware = True
    
     pOutPC.AddPoint pt1
    
    Dim v As IVector3D
     Set v = New Vector3D
    
     v.ConstructDifference pInPC.Point(1), pt1
    
    Dim pVecList As IArray
     Set pVecList = New esricore.Array
    
    Dim pClone As IClone
     Set pClone = v
     pVecList.Add pClone.Clone
    
    Dim v2 As IVector3D
     Set v2 = New Vector3D
    
    Dim currInx As Integer
     currInx = 0
     Do While (currInx < pInPC.PointCount - 1)
        Dim accumDist As Double
         accumDist = 0
         Do While (currInx < (pInPC.PointCount - 1))
             currInx = currInx + 1
            Dim pt2 As IPoint
             Set pt2 = pInPC.Point(currInx)
             v.ConstructDifference pt2, pt1
            Dim d As Double
             d = accumDist + v.Magnitude
             If (d >= step) Then
                 v2.SetComponents pt1.x, pt1.y, pt1.Z
                 If (v.Magnitude > 0) Then
                     v.Magnitude = (step - accumDist)
                     Set v2 = v2.AddVector(v)
                     pt1.x = v2.XComponent
                     pt1.y = v2.YComponent
                     pt1.Z = v2.ZComponent
                     Set pClone = pt1
                     pOutPC.AddPoint pClone.Clone
                     If (doVectors) Then
                         v2.ConstructDifference pt2, pInPC.Point(currInx - 1)
                         If (v2.Magnitude > 0) Then
                             v2.Magnitude = step * 2
                             Set pClone = v2
                             pVecList.Add pClone.Clone
                         Else
                             Err.Raise -9999, "GeomUtil:AlongPolylineZ", "Zero magnitude vector encountered"
                         End If
                     End If
                     currInx = currInx - 1
                 Else
                     Err.Raise -9999, "GeomUtil:AlongPolylineZ", "Zero magnitude vector encountered"
                 End If
                 Exit Do
             End If
             Set pClone = pt2
             Set pt1 = pClone.Clone
             accumDist = d
         Loop
     Loop
    
     If (doVectors) Then
         Set pVectors = pVecList
     End If
    
     Set AlongPolyLineZ = pOutPC
    
End Function

' Create an arrow defined by a point, which is used as its tail end, and a vector
' that defines direction and length.
'------------------------------------------------------------------------------------
Public Function Arrow( _
    ByVal pt1 As IPoint, _
    ByVal pSD As IVector3D, _
    Optional ByVal dLength As Double = 0) As IMultiPatch
    
    Dim tailRadiusFact As Double
    Dim tailLengthFact As Double
    Dim headRadiusFact As Double
    Dim stepAngle As Double
    
     tailRadiusFact = 0.1  ' size of line is multiplied by this to calculate tail radius
     tailLengthFact = 0.65 ' length of tail relative to length of input lineZ
     headRadiusFact = 1.5  ' size of tail radius is multiplied by this to calulate head _
radius
     stepAngle = 15        ' angle between circle points
    
    Dim steps As Long
     steps = (360 / stepAngle)
    
    Dim pZ As IZAware
     Set pZ = pt1
    'Debug.Print pZ.ZAware
    
     If (dLength <> 0) Then
         pSD.Magnitude = dLength
     End If
    
    Dim pt2 As IPoint
     Set pt2 = New esricore.Point
     pt2.x = pt1.x + pSD.XComponent
     pt2.y = pt1.y + pSD.YComponent
     pt2.Z = pt1.Z + pSD.ZComponent
    Dim pZAware As IZAware
     Set pZAware = pt2
     pZAware.ZAware = True
    
    Dim v As IVector3D
     Set v = New Vector3D
     v.ConstructDifference pt2, pt1
    
    Dim length As Double
     length = v.Magnitude
    
    Dim tailRadius As Double
     tailRadius = length * tailRadiusFact
    
    ' Generate vector orthogonal to ray defined by LineZ and set intensity to
    ' tailRadius length.
    Dim alt As Double
     alt = v.Inclination
     If (alt >= 0) Then
         alt = -90
     Else
         alt = 90
     End If
    
    Dim v2 As IVector3D
    Dim pClone As IClone
     Set pClone = v
     Set v2 = pClone.Clone
     v2.PolarMove 0, DegreesToRadians(alt), 0
     v2.Magnitude = tailRadius
    
    ' Resize the intensity of the vector to define front position of tailpiece
    ' (behind arrow head).
     v.Magnitude = length * tailLengthFact
    
    ' Create the arrow tail - a triangle fan that forms what looks like a
    ' hollow tube.
    Dim pTStrip As IPointCollection
     Set pTStrip = New TriangleStrip
    Dim pTransform As ITransform3D
    Dim i As Long
     For i = 0 To steps
         Set pClone = pt1
         Set pTransform = pClone.Clone
         v2.Rotate DegreesToRadians(stepAngle), v
         pTransform.MoveVector3D v2
         pTStrip.AddPoint pTransform
         Set pZ = pTransform ' a bug in 'AddPoint' has turned point zAwareness off
         pZ.ZAware = True    ' which is a problem for MoveVector3D
         Set pClone = pTransform
         Set pTransform = pClone.Clone
         pTransform.MoveVector3D v
         pTStrip.AddPoint pTransform
     Next i
    Dim pMP As IGeometryCollection
     Set pMP = New MultiPatch
     pMP.AddGeometry pTStrip
    
    ' Back of arrow head (as triangle fan) - the flat part on the rear of the cone
    Dim pFan As IPointCollection
     Set pFan = New TriangleFan
    Dim headRadius As Double
     headRadius = tailRadius * headRadiusFact
     v2.Magnitude = headRadius
    
    ' Copy first point (tail of arrow) and move to back of arrow head.
    ' This is the first point of the fan.
     Set pClone = pt1
     Set pTransform = pClone.Clone
     pTransform.MoveVector3D v
     pFan.AddPoint pTransform
     Set pZ = pTransform ' wierd ZAwareness problem again - see above in this routine
     pZ.ZAware = True
    
    ' Create circular fan of points around the first.
     Set pClone = pTransform
     For i = 0 To steps
         Set pTransform = pClone.Clone
         v2.Rotate DegreesToRadians(stepAngle), v
         pTransform.MoveVector3D v2
         pFan.AddPoint pTransform
     Next i
     pMP.AddGeometry pFan
    
    ' Arrow head - copy previous fan and redefine the first point
    ' so it becomes the tip of the arrow.
     Set pClone = pFan
     Set pFan = pClone.Clone
     Set pClone = pt2
     pFan.UpdatePoint 0, pClone.Clone
     pMP.AddGeometry pFan
    
     Set Arrow = pMP
    
End Function

' DensifyPath3D - densifies a path based on true 3d distance. Vertices
' will be spaced no greater apart than specified step distance. They can be
' spaced less. Version 1.0
'------------------------------------------------------------------------------------
Public Function DensifyPath3D(ByVal pPath As IPath, ByVal step As Double) As IPath
    
    Dim pClone As IClone
    
    Dim ptList As IPointCollection
     Set ptList = pPath
    
    Dim pt1 As IPoint
     Set pt1 = ptList.Point(0)
    
    Dim done As Boolean
     done = False
    
    Dim v As IVector3D
     Set v = New Vector3D
    
    Dim inx As Long
     inx = 1
     Do While (Not done)
        Dim pt2 As IPoint
         Set pt2 = ptList.Point(inx)
         v.ConstructDifference pt2, pt1
        Dim Dist As Double
         Dist = v.Magnitude
        Dim added As Boolean
         added = False
         If (Dist > step) Then
             Dist = Dist * 0.5      ' First see if mid point between vertices satifies
             If (Dist > step) Then  ' criteria, otherwise set to step distance. This
                 Dist = step          ' results in more even distribution of points and can
             End If                 ' can prevent some added points from being very close
             v.Magnitude = Dist     ' to others.
             Set pClone = pt1
             Set pt2 = pClone.Clone
            Dim pTransform As ITransform3D
             Set pTransform = pt2
             pTransform.MoveVector3D v
             ptList.InsertPoints inx, 1, pTransform
             added = True
         Else
             inx = inx + 1
         End If
         Set pClone = pt2
         Set pt1 = pClone.Clone
         If (inx >= ptList.PointCount) Then
             done = True
         End If
     Loop
    
     Set DensifyPath3D = ptList
    
End Function

' Current design limit - only smooths first path in polyline.
'------------------------------------------------------------------------------------
Public Function SmoothPolyline( _
    ByVal pInPolyline As IPolyline, _
    ByVal outsideIterations As Long, _
    ByVal iterations As Long, _
    ByVal stepDist As Double) As IPolyline
    
    ' Get a copy of the input polyline's first path
    Dim pGC As IGeometryCollection
     Set pGC = pInPolyline
    Dim pClone As IClone
     Set pClone = pGC.Geometry(0)
    Dim pNewPath As IPointCollection
     Set pNewPath = pClone.Clone
    
    Dim reduceStep As Double
     reduceStep = stepDist
     stepDist = stepDist * outsideIterations
    
    Dim j As Long
     For j = 1 To outsideIterations
         Set pNewPath = DensifyPath3D(pNewPath, stepDist)
        Dim x As Long
         For x = 1 To iterations
            Dim pNewerPath As IPointCollection
             Set pNewerPath = New Path
            Dim pZAware As IZAware
             Set pZAware = pNewerPath
             pZAware.ZAware = True
            Dim i As Long
             For i = 0 To (pNewPath.PointCount - 1)
                
                ' Preserve first and last points
                 If (i = 0) Then
                     pNewerPath.AddPoint pNewPath.Point(i)
                     ElseIf (i = (pNewPath.PointCount - 1)) Then
                     pNewerPath.AddPoint pNewPath.Point(i)
                     Exit For
                 End If
                
                Dim pV As IVector3D
                 Set pV = New Vector3D
                 pV.ConstructDifference pNewPath.Point(i + 1), pNewPath.Point(i)
                 If (pV.Magnitude > 0) Then
                     pV.Magnitude = pV.Magnitude * 0.5
                    Dim pNewPoint As ITransform3D
                     Set pClone = pNewPath.Point(i)
                     Set pNewPoint = pClone.Clone
                     pNewPoint.MoveVector3D pV
                    Dim Dist As Double
                     pV.ConstructDifference pNewPoint, pNewerPath.Point(pNewerPath.PointCount - 1)
                     Dist = pV.Magnitude
                     If (Dist > reduceStep) Then
                         pNewerPath.AddPoint pNewPoint
                     End If
                 Else
                    ' Without exhaustive debugging, it appears this is possible through multiple
                    ' iterations (inside/outside) that include densification and/or preservation
                    ' of 1st and last points. Since we want to prevet coincident points the
                    ' thing to do is just move to the next position.
                 End If
             Next i
             Set pClone = pNewerPath
             Set pNewPath = pClone.Clone
         Next x
         stepDist = stepDist - reduceStep
     Next j
    
    Dim pOutPolyline As IGeometryCollection
     Set pOutPolyline = New Polyline
     Set pZAware = pOutPolyline
     pZAware.ZAware = True
     pOutPolyline.AddGeometry pNewerPath
     Set SmoothPolyline = pOutPolyline
    
End Function

' 3D length of polyline with Z's, including all paths.
'------------------------------------------------------------------------------------
Public Function PolylineLength3D(pPolyline As IPolyline) As Double
    Dim pGC As IGeometryCollection
     Set pGC = pPolyline
    
    Dim pV As IVector3D
     Set pV = New Vector3D
    
    Dim Dist As Double
     Dist = 0
    
    Dim i As Long
     For i = 0 To (pGC.GeometryCount - 1)
        Dim pPC As IPointCollection
         Set pPC = pGC.Geometry(i)
        Dim j As Long
         For j = 0 To (pPC.PointCount - 2)
             pV.ConstructDifference pPC.Point(j), pPC.Point(j + 1)
             Dist = Dist + pV.Magnitude
         Next j
     Next i
    
     PolylineLength3D = Dist
End Function

' Create multipatches from LineOfSight result for alternate way of viewing
' results in 3D. All but the last two args should be results from a
' line-of-sight (LOS) calculation. The last three args are the outputs. The
' first two should be instantiated, but empty, multipatches. The last is the
' height the target would need to be raised to in order to become visible (can
' be ignored if already visible).
'
' Routine for use only in ArcGIS 8.1.2 or later.
'------------------------------------------------------------------------------------
Public Sub CreateVerticalLOSPatches( _
    bIsVis As Boolean, _
    pObsPt As IPoint, _
    pTarPt As IPoint, _
    pVisLine As IPolyline, _
    pInVisLine As IPolyline, _
    pVisPatch As IGeometryCollection, _
    pInVisPatch As IGeometryCollection, _
    dTargetHeight As Double)
    
    On Error GoTo EH
    
    Dim pGeomColl As IGeometryCollection
     Set pGeomColl = pVisLine
    
    Dim pVisMPatch As IMultiPatch
     Set pVisMPatch = pVisPatch
    
    Dim pInVisMPatch As IMultiPatch
     Set pInVisMPatch = pInVisPatch
    
     dTargetHeight = pTarPt.Z
    
    Dim i As Long
     For i = 0 To (pGeomColl.GeometryCount - 1)
        Dim pPC As IPointCollection
         Set pPC = pGeomColl.Geometry(i)
         If (i = 0) Then ' save first profile point for later
            Dim pClone As IClone
             Set pClone = pPC.Point(0)
            Dim pStartPoint As IPoint
             Set pStartPoint = pClone.Clone
         End If
         Set pClone = pPC
        Dim pVisFan As IPointCollection
         Set pVisFan = New TriangleFan
         Set pClone = pObsPt
         pVisFan.AddPoint pClone.Clone
         Set pClone = pPC
         pVisFan.AddPointCollection pClone.Clone
        
        ' Get 2D distance to last visible vertex along profile
         If (i = (pGeomColl.GeometryCount - 1)) Then
            Dim pV As IVector3D
             Set pV = New Vector3D
            Dim p1 As IPoint
             Set pClone = pObsPt
             Set p1 = pClone.Clone
             p1.Z = 0
            Dim p2 As IPoint
             Set pClone = pPC.Point(pPC.PointCount - 1)
             Set p2 = pClone.Clone
             p2.Z = 0
             pV.ConstructDifference p1, p2
            Dim dist1 As Double
             dist1 = pV.Magnitude
            
            ' save last vis point for later
            Dim pLastVisPoint As IPoint
             Set pLastVisPoint = pClone.Clone
            
             If (pInVisLine Is Nothing) Then ' bring visible fan up to target height
                 If (pTarPt.Z > pPC.Point(pPC.PointCount - 1).Z) Then
                     Set pClone = pTarPt
                     pVisFan.AddPoint pClone.Clone
                 End If
             End If
            
         End If
         pVisPatch.AddGeometry pVisFan
     Next i
    
     If (Not pInVisLine Is Nothing) Then
        
         Set pGeomColl = pInVisLine
         For i = 0 To (pGeomColl.GeometryCount - 1)
             Set pPC = pGeomColl.Geometry(i)
             Set pClone = pPC
            Dim pInVisRing As IPointCollection
             Set pInVisRing = New Ring
             Set pClone = pPC
             pInVisRing.AddPointCollection pClone.Clone
            
            ' Get 2D distance to last invisible vertex along profile
             If (i = (pGeomColl.GeometryCount - 1)) Then
                 Set pV = New Vector3D
                 Set pClone = pObsPt
                 Set p1 = pClone.Clone
                 p1.Z = 0
                 Set pClone = pPC.Point(pPC.PointCount - 1)
                 Set p2 = pClone.Clone
                 p2.Z = 0
                 pV.ConstructDifference p1, p2
                Dim dist2 As Double
                 dist2 = pV.Magnitude
                 If (dist1 < dist2) Then ' last vertex on profile belongs to invisible part
                     Set pClone = pObsPt
                     Set p1 = pClone.Clone
                     p1.Z = 0
                     Set pClone = pPC.Point(0) ' first point of last invisible part
                     Set p2 = pClone.Clone
                     p2.Z = 0
                     pV.ConstructDifference p1, p2 ' obs and first point of last part
                    Dim theDist1 As Double
                     theDist1 = pV.Magnitude
                    
                    Dim slope As Double
                     slope = (pObsPt.Z - pPC.Point(0).Z) / theDist1
                    
                     Set pClone = pPC.Point(pPC.PointCount - 1)
                    Dim pEndPoint As IPoint
                     Set pEndPoint = pClone.Clone
                    
                     Set p2 = pClone.Clone
                     p2.Z = 0
                     pV.ConstructDifference p1, p2
                    Dim theDist2 As Double
                     theDist2 = pV.Magnitude
                    
                    Dim deltaZ As Double
                     deltaZ = theDist2 * slope
                    
                    Dim theHeight As Double
                     theHeight = pObsPt.Z - deltaZ
                    
                     pEndPoint.Z = theHeight
                     Set pClone = pEndPoint
                    
                     pInVisRing.AddPoint pClone.Clone
                    
                     If (bIsVis) Then
                        ' The last profile point is not visible but that target is
                        ' because it must have a sufficient offset. Add a visible
                        ' fan to show this.
                         Set pVisFan = New TriangleFan
                         Set pClone = pObsPt
                         pVisFan.AddPoint pClone.Clone
                         Set pClone = pTarPt
                         pVisFan.AddPoint pClone.Clone
                         Set pClone = pEndPoint
                         pVisFan.AddPoint pClone.Clone
                         pVisPatch.AddGeometry pVisFan
                     Else
                         dTargetHeight = pEndPoint.Z ' it would need to be at least this
                        ' high to be seen
                     End If
                 Else
                    ' Last profile part is visible. See if target is above last
                    ' profile point (b/c of offset), and if so, add a visible
                    ' fan to show this.
                     If (bIsVis) Then ' redundant?
                         If (pTarPt.Z > pLastVisPoint.Z) Then
                             Set pVisFan = New TriangleFan
                             Set pClone = pObsPt
                             pVisFan.AddPoint pClone.Clone
                             Set pClone = pTarPt
                             pVisFan.AddPoint pClone.Clone
                             Set pClone = pLastVisPoint
                             pVisFan.AddPoint pClone.Clone
                             pVisPatch.AddGeometry pVisFan
                         End If
                     End If
                 End If
             End If
             Set pClone = pInVisRing.Point(0)
             pInVisRing.AddPoint pClone.Clone
             pInVisPatch.AddGeometry pInVisRing
             pInVisMPatch.PutRingType pInVisRing, esriMultiPatchRing
         Next i
        
     End If
    
    Exit Sub
EH:
     Err.Raise -9999, "geomUtil.CreateVerticalLOSPatches: " & Err.Source, "RUNTIME ERROR:" _
    & Err.Description
End Sub

' pFrom and pTo are start and end points for the ellipse, lenthwise. These
' define what is later referred to as the 'sight line' which is really the
' ellipse center line.
'
' dRoundness needs to be > 0 and <= 1. Small values yield thin ellipses. A value
' of 1 produces a sphere.
'
' lNumSegs defines the number of sections along sight line. The larger it is the
' smoother the resulting geometry. But it will be more expensive to render
' because of the additional geometry.
'------------------------------------------------------------------------------------
Public Function Ellipse( _
    pFrom As IPoint, _
    pTo As IPoint, _
    dRoundness As Double, _
    Optional lNumSegs As Long = 20) As IMultiPatch
    
    On Error GoTo EH
    
    ' Sight line vector
    Dim pV As IVector3D
     Set pV = New Vector3D
     pV.ConstructDifference pTo, pFrom
    
    ' Sight line length
    Dim dSightLineLength As Double
     dSightLineLength = pV.Magnitude
    
    ' Be able to copy the original From point over and over because that position
    ' is used repeatedly in a transformation
    Dim pFromClone As IClone
     Set pFromClone = pFrom
    
    ' Must be ZAware to support transform
    Dim pZAware As IZAware
     Set pZAware = pFrom
     pZAware.ZAware = True
    
    ' Initialize output multipatch
    Dim lType As Long
     lType = 1 ' fan, anything else is strip
    
    Dim pMP As IGeometryCollection
     Set pMP = New MultiPatch
    
    Dim pPrevPoints As IPointCollection
     Set pPrevPoints = New Multipoint
     pPrevPoints.AddPoint pFromClone.Clone
    
    ' Define vector that is orthogonal to sight line vector. This ends up being
    ' rotated around the sight line like spokes around the hub of a wheel, with
    ' multipatch vertices generated at each step.
    Dim pSpokeV As IVector3D
     Set pSpokeV = New Vector3D
    Dim c As Double
     c = Sqr((pV.XComponent * pV.XComponent) + (pV.YComponent * pV.YComponent))
     pSpokeV.XComponent = pV.XComponent * (-pV.ZComponent / c)
     pSpokeV.YComponent = pV.YComponent * (-pV.ZComponent / c)
     pSpokeV.ZComponent = c
    
    ' Stuff for smooth shading
    Dim pEProps As IEncode3DProperties
     Set pEProps = New GeometryEnvironment
    
    ' Initial shading vector goes in the opposite direction as the sight line
    Dim pNormalV As IVector3D
     Set pNormalV = New Vector3D
     pNormalV.XComponent = pV.XComponent * -1
     pNormalV.YComponent = pV.YComponent * -1
     pNormalV.ZComponent = pV.ZComponent * -1
     pNormalV.Normalize
    
    Dim m As Double
     pEProps.PackNormal pNormalV, m
     pPrevPoints.Point(0).m = m
    
    Dim pSpokeVClone As IClone
     Set pSpokeVClone = pSpokeV
    
    ' Vector used to get diameter of ellipse at different positions along its
    ' length.
    Dim pEV As IVector3D
     Set pEV = New Vector3D
     pEV.XComponent = 1
     pEV.YComponent = 0
     pEV.ZComponent = 0
     pEV.Azimuth = DegreesToRadians(180)
    
    ' Translate number of desired segments into the angle increment needed to
    ' accomplish that in the loop below.
    Dim dAngleInc As Double
     dAngleInc = DegreesToRadians(180 / lNumSegs) * -1
    
    ' Factors used defining ellipse
    Dim a As Double
     a = 1
    Dim b As Double
     b = dRoundness
    
    Dim i As Long
     For i = 1 To lNumSegs
        
        ' Increment the vector represeting the ellipse.
         pEV.Azimuth = pEV.Azimuth + dAngleInc
        Dim x As Double
         x = a * Cos(pEV.Azimuth)
        Dim y As Double
         y = b * Sin(pEV.Azimuth)
        
        ' Distance along length of sight line for current position. This is the hub
        ' that spokes will be generated around. The magnitude is translated into
        ' world coordinate units here.
         pV.Magnitude = dSightLineLength * ((x + 1) * 0.5)
        
        ' Radius of ellipse at current hub position. The magnitude is translated
        ' into world coordinate units here.
         pSpokeV.Magnitude = (dSightLineLength * 0.5) * y
        
        ' Generate point at hub position on sight line. World coordinates.
        Dim pHubTrans As ITransform3D
         Set pHubTrans = pFromClone.Clone
         pHubTrans.MoveVector3D pV
        
        Dim pHubClone As IClone ' used in loop later
         Set pHubClone = pHubTrans
        
        Dim pNewPoints As IPointCollection
         Set pNewPoints = New Multipoint
        
        Dim dRadMax As Double
         dRadMax = geomutil.DegreesToRadians(360)
        
        Dim lSegsAround As Long
         lSegsAround = (lNumSegs * 2) * dRoundness
         If (lSegsAround < 10) Then
             lSegsAround = 10
         End If
        
        Dim dRotateStep As Double
         dRotateStep = dRadMax / lSegsAround
        
        '
        ' Calculate point along sight line, relative to hub point, that's used
        ' as orign point for lighting vector. Where the ellipse is thick, near
        ' the center (max y), there is little offset. Towards the ends of the
        ' ellipse, where it starts getting thinner (min y), there is greater and
        ' greater offset. Basically, the offset for x is a function of y. This
        ' works for all levels of roundness from thin ellipses to circles.
        '
        Dim pVClone As IClone
         Set pVClone = pV
         Set pNormalV = pVClone.Clone
        
        Dim dX_Adjusted As Double
         dX_Adjusted = x * (1 - (dRoundness - y))
        
         pNormalV.Magnitude = dSightLineLength * ((dX_Adjusted + 1) * 0.5)
        
        Dim pNormalOrigin As IPoint
         Set pNormalOrigin = pFromClone.Clone
        Dim pNormalTrans As ITransform3D
         Set pNormalTrans = pNormalOrigin
         pNormalTrans.MoveVector3D pNormalV
        
        ' Rotate around the sight line generating multipatch vertices
        Dim j As Double
         For j = 0 To lSegsAround
            ' copy the current hub point and  transform it outward along the spoke
            Dim pSpokeTrans As ITransform3D
             Set pSpokeTrans = pHubClone.Clone
             pSpokeTrans.MoveVector3D pSpokeV
            
            ' construct normal and encode
             pNormalV.ConstructDifference pSpokeTrans, pNormalOrigin
             pNormalV.Normalize
             pEProps.PackNormal pNormalV, m
            Dim pSpokePoint As IPoint
             Set pSpokePoint = pSpokeTrans
             pSpokePoint.m = m
            
            ' add vertex
             pNewPoints.AddPoint pSpokePoint
            
            ' increment
             pSpokeV.Rotate dRotateStep, pV
         Next j
        
         If (lType = 1) Then ' this is the first part of the multipatch, a fan
            Dim pPart As IPointCollection
             Set pPart = New TriangleFan
             pPart.AddPointCollection pPrevPoints
             pPart.AddPointCollection pNewPoints
             lType = 2
         Else                ' subsequent parts are strips
             Set pPart = New TriangleStrip
            ' connect vertices from this interation to those from previous
            ' interweave them to make the strips
             For j = 0 To (pNewPoints.PointCount - 1)
                 pPart.AddPoint pPrevPoints.Point(j)
                 pPart.AddPoint pNewPoints.Point(j)
             Next j
         End If
        
         pMP.AddGeometry pPart
        
        ' Save vertices for reuse in next iteration
        Dim pClone As IClone
         Set pClone = pNewPoints
         Set pPrevPoints = pClone.Clone
        
     Next i
    
    ' Last point at end of multipatch
     Set pHubTrans = pFromClone.Clone
     pV.Magnitude = dSightLineLength
     pHubTrans.MoveVector3D pV
    
    ' Normal for shading - at end point for patch equals sight line vector
     pV.Normalize
     pEProps.PackNormal pV, m
    Dim pHubPoint As IPoint
     Set pHubPoint = pHubTrans
     pHubPoint.m = m
    
    ' A fan as the last part closes and ends the patch
     Set pPart = New TriangleFan
     pPart.AddPoint pHubPoint
    Dim lCnt As Long
     lCnt = pPrevPoints.PointCount
     For i = (lCnt - 1) To 0 Step -1
         pPart.AddPoint pPrevPoints.Point(i)
     Next i
     pMP.AddGeometry pPart
    
    ' MAwareness for shading
    Dim pMAware As IMAware
     Set pMAware = pMP
     pMAware.MAware = True
    
     Set Ellipse = pMP
    
    Exit Function
EH:
     Err.Raise Err.Number, Err.Source, Err.Description & " in GeomUtil routine Ellipse"
End Function




