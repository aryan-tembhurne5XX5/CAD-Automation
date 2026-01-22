' ==========================================
' PART EXPORT — ASSEMBLY READY (FINAL)
' ==========================================

Imports System.Text
Imports System.IO

Dim partDoc As PartDocument = ThisApplication.ActiveDocument
Dim cd As PartComponentDefinition = partDoc.ComponentDefinition

Dim sb As New StringBuilder()

sb.AppendLine("{")
sb.AppendLine("""part_name"": """ & partDoc.DisplayName & """,")

' -------------------------
' iProperties
' -------------------------
Try
    Dim props = partDoc.PropertySets.Item("Design Tracking Properties")
    sb.AppendLine("""part_number"": """ & props.Item("Part Number").Value & """,")
    sb.AppendLine("""description"": """ & props.Item("Description").Value & """,")
Catch
    sb.AppendLine("""part_number"": null,")
    sb.AppendLine("""description"": null,")
End Try

' -------------------------
' HOLES (iLogic SAFE)
' -------------------------
sb.AppendLine("""holes"": [")

Dim firstHole As Boolean = True

For Each h As HoleFeature In cd.Features.HoleFeatures
    If h.Suppressed Then Continue For

    Dim pdef = h.PlacementDefinition
    If pdef.Type <> 0 Then Continue For ' must be sketch-based

    ' Axis = sketch plane normal
    Dim plane = pdef.Sketch.PlanarEntityGeometry
    Dim axisX As Double = plane.Normal.X
    Dim axisY As Double = plane.Normal.Y
    Dim axisZ As Double = plane.Normal.Z

    ' Diameter (cm → mm)
    Dim diaMM As Double = h.HoleDiameter.Value * 10

    ' One hole feature can generate multiple holes (patterns)
    For Each sp As SketchPoint In pdef.HoleCenterPoints
        If Not firstHole Then sb.AppendLine(",")
        firstHole = False

        Dim p = sp.Geometry3d

        sb.AppendLine("{")
        sb.AppendLine("""feature_name"": """ & h.Name & """,")
        sb.AppendLine("""hole_type"": " & h.HoleType & ",")
        sb.AppendLine("""diameter_mm"": " & Math.Round(diaMM, 3) & ",")
        sb.AppendLine("""center"": [" &
            p.X & "," & p.Y & "," & p.Z & "],")
        sb.AppendLine("""axis"": [" &
            axisX & "," & axisY & "," & axisZ & "],")
        sb.AppendLine("""threaded"": " & h.Tapped.ToString().ToLower())
        sb.Append("}")
    Next
Next

sb.AppendLine("]")
sb.AppendLine("}")

' -------------------------
' WRITE FILE
' -------------------------
Dim outPath As String =
    System.IO.Path.Combine(
        System.IO.Path.GetDirectoryName(partDoc.FullFileName),
        partDoc.DisplayName & "_assembly_data.json"
    )

System.IO.File.WriteAllText(outPath, sb.ToString())

MessageBox.Show("✅ Part data exported successfully:" & vbLf & outPath)