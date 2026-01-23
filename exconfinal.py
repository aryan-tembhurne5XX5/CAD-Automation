Imports System.IO
Imports System.Text
Imports Inventor

Sub Main()
    ' --- 1. Environment Check ---
    Dim oDoc As Document = ThisDoc.Document
    If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("Run this rule inside a Part (.ipt) file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return
    End If

    Dim oPartDoc As PartDocument = CType(oDoc, PartDocument)
    Dim oDef As PartComponentDefinition = oPartDoc.ComponentDefinition

    ' --- 2. Initialize JSON Builder ---
    Dim sb As New StringBuilder()
    sb.AppendLine("{")

    ' ==========================================
    ' SECTION: PART METADATA
    ' ==========================================
    Dim partName As String = System.IO.Path.GetFileNameWithoutExtension(oPartDoc.FullFileName)
    sb.AppendLine("  ""part_metadata"": {")
    sb.AppendLine("    ""file_name"": """ & EscapeJson(partName) & """,")
    sb.AppendLine("    ""full_path"": """ & EscapeJson(oPartDoc.FullFileName) & """,")
    sb.AppendLine("    ""part_number"": """ & GetProp(oDoc, "Design Tracking Properties", "Part Number") & """,")
    sb.AppendLine("    ""description"": """ & GetProp(oDoc, "Design Tracking Properties", "Description") & """,")
    sb.AppendLine("    ""material"": """ & GetProp(oDoc, "Design Tracking Properties", "Material") & """,")
    sb.AppendLine("    ""units"": ""mm"",")
    Try
        sb.AppendLine("    ""mass_kg"": " & oDef.MassProperties.Mass.ToString(System.Globalization.CultureInfo.InvariantCulture))
    Catch
        sb.AppendLine("    ""mass_kg"": 0")
    End Try
    sb.AppendLine("  },")

    ' ==========================================
    ' SECTION: COORDINATE SYSTEM (Identity)
    ' ==========================================
    sb.AppendLine("  ""coordinate_system"": {")
    sb.AppendLine("    ""origin_mm"": { ""x"": 0, ""y"": 0, ""z"": 0 },")
    sb.AppendLine("    ""x_axis"": { ""x"": 1, ""y"": 0, ""z"": 0 },")
    sb.AppendLine("    ""y_axis"": { ""x"": 0, ""y"": 1, ""z"": 0 },")
    sb.AppendLine("    ""z_axis"": { ""x"": 0, ""y"": 0, ""z"": 1 }")
    sb.AppendLine("  },")

    ' ==========================================
    ' SECTION: BOUNDING BOX
    ' ==========================================
    Dim minX As Double = 0 : Dim minY As Double = 0 : Dim minZ As Double = 0
    Dim maxX As Double = 0 : Dim maxY As Double = 0 : Dim maxZ As Double = 0
    Try
        Dim rb As Box = oDef.RangeBox
        minX = rb.MinPoint.X * 10.0 : minY = rb.MinPoint.Y * 10.0 : minZ = rb.MinPoint.Z * 10.0
        maxX = rb.MaxPoint.X * 10.0 : maxY = rb.MaxPoint.Y * 10.0 : maxZ = rb.MaxPoint.Z * 10.0
    Catch
    End Try

    sb.AppendLine("  ""bounding_box_mm"": {")
    sb.AppendLine("    ""min"": { ""x"": " & Num(minX) & ", ""y"": " & Num(minY) & ", ""z"": " & Num(minZ) & " },")
    sb.AppendLine("    ""max"": { ""x"": " & Num(maxX) & ", ""y"": " & Num(maxY) & ", ""z"": " & Num(maxZ) & " }")
    sb.AppendLine("  },")

    ' ==========================================
    ' SECTION: CONNECTION POINTS (Holes)
    ' ==========================================
    sb.AppendLine("  ""connection_points"": [")
    Dim connList As New List(Of String)
    
    For Each oHole As HoleFeature In oDef.Features.HoleFeatures
        If oHole.Suppressed Then Continue For

        ' Check Sketch Placement
        If oHole.PlacementDefinition.Type = ObjectTypeEnum.kSketchHolePlacementDefinitionObject Then
            Dim oSketchPlacement As SketchHolePlacementDefinition = oHole.PlacementDefinition
            
            ' -- Extract Hole Properties --
            Dim holeTypeStr As String = "Simple"
            Dim diam As Double = 0.0
            Dim minorDiam As String = "null"
            Dim pitch As String = "null"
            Dim isThreaded As Boolean = oHole.Tapped
            Dim isThrough As Boolean = (oHole.ExtentType = PartFeatureExtentTypeEnum.kThroughAllExtent)
            Dim depthStr As String = "null"
            Dim angleStr As String = "null"

            ' Diameter
            Try
                diam = oHole.HoleDiameter.Value * 10.0
            Catch
            End Try

            ' Specific Type Handling
            If oHole.HoleType = HoleTypeEnum.kCounterBoreHole Then holeTypeStr = "Counterbore"
            If oHole.HoleType = HoleTypeEnum.kCounterSinkHole Then 
                holeTypeStr = "Countersink"
                Try
                   ' angleStr = (oHole.CsinkAngle.Value * (180/Math.PI)).ToString() ' If needed
                Catch
                End Try
            End If
            If isThreaded Then holeTypeStr = "Tapped"

            ' Thread Info
            If isThreaded Then
                Try
                    Dim tapInfo As HoleTapInfo = oHole.TapInfo
                    ' Estimate minor diam (approx) or try to read from TapInfo if available
                    ' API access to MinorDiameter is limited, usually derived from ThreadInfo
                    ' pitch = tapInfo.Pitch.Value * 10 ' often fails on standard threads
                Catch
                End Try
            End If

            ' Depth
            If Not isThrough Then
               Try
                   ' This is tricky, depends on Extent object type
                   Dim distExt As DistanceExtent = CType(oHole.Extent, DistanceExtent)
                   depthStr = Num(distExt.Distance.Value * 10.0)
               Catch
               End Try
            End If

            ' Pattern Detection (Sketch based pattern)
            Dim isPatterned As Boolean = (oSketchPlacement.HoleCenterPoints.Count > 1)
            Dim patternType As String = If(isPatterned, "SketchPattern", "None")

            ' -- Iterate Points --
            Dim ptIndex As Integer = 0
            For Each oPoint As SketchPoint In oSketchPlacement.HoleCenterPoints
                ptIndex += 1
                Dim cpSb As New StringBuilder()
                
                ' Geometry
                Dim pX As Double = oPoint.Geometry3d.X * 10.0
                Dim pY As Double = oPoint.Geometry3d.Y * 10.0
                Dim pZ As Double = oPoint.Geometry3d.Z * 10.0
                
                Dim oSketch As PlanarSketch = CType(oPoint.Parent, PlanarSketch)
                Dim nX As Double = oSketch.PlanarEntityGeometry.Normal.X
                Dim nY As Double = oSketch.PlanarEntityGeometry.Normal.Y
                Dim nZ As Double = oSketch.PlanarEntityGeometry.Normal.Z

                ' Entry Face ID
                Dim entryFaceId As String = "null"
                Try
                    ' Attempt to get internal name of the sketch plane face
                    ' entryFaceId = """" & EscapeJson(oSketch.PlanarEntity.InternalName) & """"
                Catch
                End Try

                cpSb.AppendLine("    {")
                cpSb.AppendLine("      ""id"": """ & System.Guid.NewGuid().ToString() & """,")
                cpSb.AppendLine("      ""feature_name"": """ & EscapeJson(oHole.Name) & """,")
                cpSb.AppendLine("      ""feature_type"": ""Hole"",")
                
                cpSb.AppendLine("      ""hole_properties"": {")
                cpSb.AppendLine("        ""hole_type"": """ & holeTypeStr & """,")
                cpSb.AppendLine("        ""diameter_mm"": " & Num(diam) & ",")
                cpSb.AppendLine("        ""minor_diameter_mm"": " & minorDiam & ",")
                cpSb.AppendLine("        ""thread_pitch_mm"": " & pitch & ",")
                cpSb.AppendLine("        ""is_threaded"": " & isThreaded.ToString().ToLower() & ",")
                cpSb.AppendLine("        ""is_through"": " & isThrough.ToString().ToLower() & ",")
                cpSb.AppendLine("        ""depth_mm"": " & depthStr & ",")
                cpSb.AppendLine("        ""countersink_angle_deg"": " & angleStr)
                cpSb.AppendLine("      },")

                cpSb.AppendLine("      ""geometry"": {")
                cpSb.AppendLine("        ""center_mm"": { ""x"": " & Num(pX) & ", ""y"": " & Num(pY) & ", ""z"": " & Num(pZ) & " },")
                cpSb.AppendLine("        ""axis"": { ""x"": " & Num(nX) & ", ""y"": " & Num(nY) & ", ""z"": " & Num(nZ) & " },")
                cpSb.AppendLine("        ""entry_face_normal"": { ""x"": " & Num(nX) & ", ""y"": " & Num(nY) & ", ""z"": " & Num(nZ) & " }")
                cpSb.AppendLine("      },")

                cpSb.AppendLine("      ""supporting_faces"": {")
                cpSb.AppendLine("        ""entry_face_id"": " & entryFaceId & ", ""exit_face_id"": null")
                cpSb.AppendLine("      },")
                
                cpSb.AppendLine("      ""thickness_along_axis_mm"": null,")

                cpSb.AppendLine("      ""pattern_info"": {")
                cpSb.AppendLine("        ""is_patterned"": " & isPatterned.ToString().ToLower() & ",")
                cpSb.AppendLine("        ""pattern_type"": """ & patternType & """,")
                cpSb.AppendLine("        ""pattern_parent"": """ & EscapeJson(oHole.Name) & """,")
                cpSb.AppendLine("        ""pattern_index"": " & ptIndex)
                cpSb.AppendLine("      },")

                cpSb.AppendLine("      ""confidence"": 1.0")
                cpSb.Append("    }")
                
                connList.Add(cpSb.ToString())
            Next
        End If
    Next
    sb.Append(String.Join("," & vbCrLf, connList.ToArray()))
    sb.AppendLine()
    sb.AppendLine("  ],")

    ' ==========================================
    ' SECTION: FACES (B-Rep Analysis)
    ' ==========================================
    sb.AppendLine("  ""faces"": [")
    Dim faceList As New List(Of String)
    Try
        ' Limit face count for performance on huge parts
        Dim faces As Faces = oDef.SurfaceBodies(1).Faces
        Dim maxFaces As Integer = 500
        Dim count As Integer = 0
        
        For Each f As Face In faces
            If count > maxFaces Then Exit For
            count += 1
            
            Dim fType As String = "Other"
            If f.SurfaceType = SurfaceTypeEnum.kPlanarSurface Then fType = "Planar"
            If f.SurfaceType = SurfaceTypeEnum.kCylinderSurface Then fType = "Cylindrical"
            If f.SurfaceType = SurfaceTypeEnum.kConeSurface Then fType = "Conical"
            
            Dim area As Double = f.Evaluator.Area * 100.0 ' cm2 -> mm2
            Dim pt As Point = f.PointOnFace
            Dim cenX As Double = pt.X * 10.0
            Dim cenY As Double = pt.Y * 10.0
            Dim cenZ As Double = pt.Z * 10.0
            
            ' Normal at Point
            Dim n(2) As Double
            Dim p(2) As Double
            p(0) = pt.X : p(1) = pt.Y : p(2) = pt.Z
            f.Evaluator.GetNormalAtPoint(p, n)
            
            Dim fSb As New StringBuilder()
            fSb.AppendLine("    {")
            fSb.AppendLine("      ""face_id"": """ & count & """,") ' InternalName is too long, using Index
            fSb.AppendLine("      ""face_type"": """ & fType & """,")
            fSb.AppendLine("      ""area_mm2"": " & Num(area) & ",")
            fSb.AppendLine("      ""normal"": { ""x"": " & Num(n(0)) & ", ""y"": " & Num(n(1)) & ", ""z"": " & Num(n(2)) & " },")
            fSb.AppendLine("      ""center_mm"": { ""x"": " & Num(cenX) & ", ""y"": " & Num(cenY) & ", ""z"": " & Num(cenZ) & " }")
            fSb.Append("    }")
            faceList.Add(fSb.ToString())
        Next
    Catch
    End Try
    sb.Append(String.Join("," & vbCrLf, faceList.ToArray()))
    sb.AppendLine()
    sb.AppendLine("  ],")

    ' ==========================================
    ' SECTION: FEATURE GRAPH
    ' ==========================================
    sb.AppendLine("  ""feature_graph"": {")
    sb.AppendLine("    ""features"": [")
    Dim featGraphList As New List(Of String)
    
    For Each feat As PartFeature In oDef.Features
        If feat.Suppressed Then Continue For
        
        Dim fType As String = "Other"
        If TypeOf feat Is HoleFeature Then fType = "Hole"
        If TypeOf feat Is ExtrudeFeature Then fType = "Extrude"
        If TypeOf feat Is CutFeature Then fType = "Cut"
        If TypeOf feat Is RectangularPatternFeature Then fType = "Pattern"
        If TypeOf feat Is CircularPatternFeature Then fType = "Pattern"

        ' Dependencies
        Dim childList As New List(Of String)
        Try
            For Each child As PartFeature In feat.DependentFeatures
                childList.Add("""" & EscapeJson(child.Name) & """")
            Next
        Catch
        End Try

        Dim fgSb As New StringBuilder()
        fgSb.AppendLine("      {")
        fgSb.AppendLine("        ""feature_name"": """ & EscapeJson(feat.Name) & """,")
        fgSb.AppendLine("        ""feature_type"": """ & fType & """,")
        fgSb.AppendLine("        ""parent_features"": [],") ' Parents hard to traverse up via API generically
        fgSb.AppendLine("        ""child_features"": [" & String.Join(",", childList.ToArray()) & "]")
        fgSb.Append("      }")
        featGraphList.Add(fgSb.ToString())
    Next
    sb.Append(String.Join("," & vbCrLf, featGraphList.ToArray()))
    sb.AppendLine()
    sb.AppendLine("    ]")
    sb.AppendLine("  }")

    sb.AppendLine("}") ' END ROOT

    ' --- Write File ---
    Dim jsonPath As String = System.IO.Path.ChangeExtension(oPartDoc.FullFileName, "json")
    System.IO.File.WriteAllText(jsonPath, sb.ToString())
    MessageBox.Show("Deep Data Export Complete: " & vbCrLf & jsonPath, "Success")
End Sub

' Helper: Format Number
Function Num(val As Double) As String
    Return val.ToString(System.Globalization.CultureInfo.InvariantCulture)
End Function

' Helper: Escape JSON
Function EscapeJson(str As String) As String
    If str Is Nothing Then Return ""
    Return str.Replace("\", "\\").Replace("""", "\""").Replace(vbCrLf, "")
End Function

' Helper: Get Property
Function GetProp(doc As Document, setName As String, propName As String) As String
    Try
        Return EscapeJson(doc.PropertySets.Item(setName).Item(propName).Value.ToString())
    Catch
        Return ""
    End Try
End Function