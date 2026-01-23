Imports System.IO
Imports System.Text
Imports Inventor

Sub Main()
    ' --- 1. Environment Validation ---
    Dim oDoc As Document = ThisDoc.Document

    If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("This rule can only be run inside a Part file (.ipt).", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return
    End If

    Dim oPartDoc As PartDocument = CType(oDoc, PartDocument)
    Dim oDef As PartComponentDefinition = oPartDoc.ComponentDefinition

    ' --- 2. JSON Construction Setup ---
    Dim sb As New StringBuilder()
    sb.AppendLine("{")

    ' --- 3. General Info ---
    Dim partName As String = System.IO.Path.GetFileNameWithoutExtension(oPartDoc.FullFileName)
    sb.AppendLine("  ""fileName"": """ & EscapeJson(partName) & """,")
    
    ' Physical Properties
    Try
        Dim mass As Double = oPartDoc.ComponentDefinition.MassProperties.Mass
        sb.AppendLine("  ""mass_kg"": " & mass.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
    Catch
        sb.AppendLine("  ""mass_kg"": 0,")
    End Try

    ' --- 4. DETAILED HOLE DATA ---
    sb.AppendLine("  ""connectionPoints"": [")
    
    Dim connectionList As New List(Of String)

    ' Iterate through standard HoleFeatures
    For Each oHole As HoleFeature In oDef.Features.HoleFeatures
        If oHole.Suppressed Then Continue For

        ' Get the placement definition
        Dim oPlacement As HolePlacementDefinition = oHole.PlacementDefinition

        ' FIX: Correctly check type using ObjectTypeEnum
        ' kSketchHolePlacementDefinitionObject = 83912192
        If oPlacement.Type = ObjectTypeEnum.kSketchHolePlacementDefinitionObject Then
            
            ' FIX: Cast to SketchHolePlacementDefinition to access specific properties
            Dim oSketchPlacement As SketchHolePlacementDefinition
            oSketchPlacement = oPlacement

            Dim isThreaded As Boolean = oHole.Tapped
            Dim holeType As String = If(isThreaded, "Threaded", "Simple")
            Dim diameterMm As Double = 0.0

            Try
                ' Internal units are cm, convert to mm (* 10)
                diameterMm = oHole.HoleDiameter.Value * 10.0
            Catch
                diameterMm = 0.0
            End Try

            ' FIX: The correct API property is HoleCenterPoints (Not SketchCenterPoints)
            Dim oSketchPoints As ObjectCollection = oSketchPlacement.HoleCenterPoints

            For Each oPoint As SketchPoint In oSketchPoints
                
                ' 3D Center Calculation (Model Space)
                Dim pX As Double = oPoint.Geometry3d.X * 10.0
                Dim pY As Double = oPoint.Geometry3d.Y * 10.0
                Dim pZ As Double = oPoint.Geometry3d.Z * 10.0

                ' Axis Vector Calculation (Sketch Plane Normal)
                Dim oSketch As PlanarSketch = CType(oPoint.Parent, PlanarSketch)
                Dim oPlane As Plane = oSketch.PlanarEntityGeometry
                
                Dim nX As Double = oPlane.Normal.X
                Dim nY As Double = oPlane.Normal.Y
                Dim nZ As Double = oPlane.Normal.Z

                ' Build JSON Object
                Dim cb As New StringBuilder()
                cb.AppendLine("    {")
                cb.AppendLine("      ""featureName"": """ & EscapeJson(oHole.Name) & """,")
                cb.AppendLine("      ""type"": """ & holeType & """,")
                cb.AppendLine("      ""diameterMm"": " & diameterMm.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
                cb.AppendLine("      ""center"": { ""x"": " & pX.ToString(System.Globalization.CultureInfo.InvariantCulture) & ", ""y"": " & pY.ToString(System.Globalization.CultureInfo.InvariantCulture) & ", ""z"": " & pZ.ToString(System.Globalization.CultureInfo.InvariantCulture) & " },")
                cb.AppendLine("      ""axis"": { ""x"": " & nX.ToString(System.Globalization.CultureInfo.InvariantCulture) & ", ""y"": " & nY.ToString(System.Globalization.CultureInfo.InvariantCulture) & ", ""z"": " & nZ.ToString(System.Globalization.CultureInfo.InvariantCulture) & " }")
                cb.Append("    }")
                connectionList.Add(cb.ToString())
            Next
        End If
    Next

    sb.Append(String.Join("," & vbCrLf, connectionList.ToArray()))
    sb.AppendLine()
    sb.AppendLine("  ]")
    sb.AppendLine("}") 

    ' --- 5. File Output ---
    Dim jsonPath As String = System.IO.Path.ChangeExtension(oPartDoc.FullFileName, "json")
    
    If String.IsNullOrEmpty(oPartDoc.FullFileName) Then
        MessageBox.Show("Please save the IPT file before exporting.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Exit Sub
    End If

    Try
        System.IO.File.WriteAllText(jsonPath, sb.ToString())
        MessageBox.Show("Export successful!" & vbCrLf & jsonPath, "iLogic Export")
    Catch ex As Exception
        MessageBox.Show("Error writing file: " & ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

End Sub

' Helper: Escape special chars for JSON
Function EscapeJson(str As String) As String
    If str Is Nothing Then Return ""
    Return str.Replace("\", "\\").Replace("""", "\""").Replace(vbCrLf, "\n")
End Function