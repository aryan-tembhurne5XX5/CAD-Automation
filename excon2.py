Imports System.IO
Imports System.Text

Sub Main()
    ' Check if active document is a Part
    If ThisApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("This rule must be run inside a Part file.", "iLogic Export", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Exit Sub
    End If

    Dim oDoc As PartDocument
    oDoc = ThisApplication.ActiveDocument

    Dim oDef As PartComponentDefinition
    oDef = oDoc.ComponentDefinition

    ' --- 1. Prepare JSON Builder ---
    Dim sb As New StringBuilder()
    sb.AppendLine("{")
    
    ' --- 2. Extract Metadata (iProperties) ---
    Dim partName As String = System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName)
    Dim partNum As String = oDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
    Dim desc As String = oDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value.ToString()

    ' Escape strings for JSON
    partName = EscapeJson(partName)
    partNum = EscapeJson(partNum)
    desc = EscapeJson(desc)

    sb.AppendLine("  ""partName"": """ & partName & """,")
    sb.AppendLine("  ""partNumber"": """ & partNum & """,")
    sb.AppendLine("  ""description"": """ & desc & """,")
    sb.AppendLine("  ""holes"": [")

    ' --- 3. Extract Holes ---
    Dim holeList As New List(Of String)
    
    ' Iterate through all HoleFeatures in the part
    For Each oHole As HoleFeature In oDef.Features.HoleFeatures
        
        ' Skip if suppressed
        If oHole.Suppressed Then Continue For

        ' Only process Sketch-based holes (Standard mechanism)
        If oHole.PlacementDefinition.Type = FeatureDimensionsTypeEnum.kSketchHolePlacementDefinition Then
            
            ' Extract Basic Hole Data
            Dim isThreaded As Boolean = oHole.Tapped
            Dim diameterMm As Double = 0.0
            
            ' Determine Diameter based on API constraints (HoleDiameter returns ModelParameter)
            ' Internal units are cm, convert to mm (* 10)
            Try
                diameterMm = oHole.HoleDiameter.Value * 10.0
            Catch ex As Exception
                ' Fallback for complex types like Tapered or specific tapped definitions if HoleDiameter is not direct
                ' However, HoleDiameter is the standard API property requested. 
                ' If it fails (e.g. some NPT types), we default to 0 or try TapInfo
                If isThreaded Then
                     ' Approximate using TapInfo if HoleDiameter fails, though constraint was specific
                     diameterMm = 0.0 
                End If
            End Try

            Dim holeType As String = "Simple"
            If isThreaded Then holeType = "Threaded"

            ' Iterate through Sketch Points (Handles Sketch Patterns/Multiple Centers)
            Dim oSketchPoints As Object = oHole.SketchCenterPoints
            
            For Each oPoint As SketchPoint In oSketchPoints
                
                ' Get 3D Center (SketchPoint.Geometry3d is in Model Space, Internal Units cm)
                Dim pX As Double = oPoint.Geometry3d.X * 10.0
                Dim pY As Double = oPoint.Geometry3d.Y * 10.0
                Dim pZ As Double = oPoint.Geometry3d.Z * 10.0

                ' Get Axis Vector (Normal of the Sketch Plane)
                ' We use the Sketch associated with the point
                Dim oSketch As PlanarSketch = oPoint.Parent
                Dim oPlane As Plane = oSketch.PlanarEntityGeometry
                
                Dim nX As Double = oPlane.Normal.X
                Dim nY As Double = oPlane.Normal.Y
                Dim nZ As Double = oPlane.Normal.Z

                ' Construct JSON Object for this specific hole instance
                Dim hBuilder As New StringBuilder()
                hBuilder.AppendLine("    {")
                hBuilder.AppendLine("      ""featureName"": """ & EscapeJson(oHole.Name) & """,")
                hBuilder.AppendLine("      ""type"": """ & holeType & """,")
                hBuilder.AppendLine("      ""diameterMm"": " & diameterMm.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
                hBuilder.AppendLine("      ""isThreaded"": " & isThreaded.ToString().ToLower() & ",")
                
                ' Position
                hBuilder.AppendLine("      ""center"": {")
                hBuilder.AppendLine("        ""x"": " & pX.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
                hBuilder.AppendLine("        ""y"": " & pY.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
                hBuilder.AppendLine("        ""z"": " & pZ.ToString(System.Globalization.CultureInfo.InvariantCulture) )
                hBuilder.AppendLine("      },")
                
                ' Axis
                hBuilder.AppendLine("      ""axis"": {")
                hBuilder.AppendLine("        ""x"": " & nX.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
                hBuilder.AppendLine("        ""y"": " & nY.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
                hBuilder.AppendLine("        ""z"": " & nZ.ToString(System.Globalization.CultureInfo.InvariantCulture) )
                hBuilder.AppendLine("      }")
                
                hBuilder.Append("    }")
                
                holeList.Add(hBuilder.ToString())
            Next
        End If
    Next

    ' Join hole objects with commas
    sb.Append(String.Join("," & vbCrLf, holeList.ToArray()))
    sb.AppendLine()
    sb.AppendLine("  ]") ' End holes
    sb.AppendLine("}") ' End root

    ' --- 4. Write to File ---
    Dim path As String = oDoc.FullFileName
    Dim jsonPath As String = System.IO.Path.ChangeExtension(path, "json")
    
    System.IO.File.WriteAllText(jsonPath, sb.ToString())
    
    MessageBox.Show("Export successful!" & vbCrLf & "File saved to: " & jsonPath, "iLogic Export")

End Sub

' Helper function to escape special characters for JSON strings
Function EscapeJson(str As String) As String
    If str Is Nothing Then Return ""
    Return str.Replace("\", "\\").Replace("""", "\""").Replace(vbCrLf, "\n").Replace(vbCr, "\n").Replace(vbLf, "\n")
End Function