' ==========================================
' PART EXPORT — ASSEMBLY READY (FIXED)
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
' HOLES (CORRECT METHOD)
' -------------------------
sb.AppendLine("""holes"": [")

Dim firstHole As Boolean = True

For Each h As HoleFeature In cd.Features.HoleFeatures
    If h.Suppressed Then Continue For

    Dim pdef = h.PlacementDefinition
    Dim hdef = h.HoleDefinition

    ' Hole axis = sketch plane normal
    Dim axisVec(2) As Double
    If pdef.Type = 0 Then ' kSketchPlacementDefinition
        Dim plane = pdef.Sketch.PlanarEntityGeometry
        axisVec(0) = plane.Normal.X
        axisVec(1) = plane.Normal.Y
        axisVec(2) = plane.Normal.Z
    Else
        Continue For
    End If

    ' Diameter (cm → mm)
    Dim diaMM As Double = 0
    Try
        diaMM = hdef.Diameter.Value * 10
    Catch
    End Try

    ' Each hole may have multiple centers (patterns)
    For Each sp As SketchPoint In pdef.HoleCenterPoints
        If Not firstHole Then sb.AppendLine(",")
        firstHole = False

        Dim p = sp.Geometry3d

        sb.AppendLine("{")
        sb.AppendLine("""feature_name"": """ & h.Name & """,")
        sb.AppendLine("""diameter_mm"": " & Math.Round(diaMM, 3) & ",")

        sb.AppendLine("""center"": [" &
            p.X & "," & p.Y & "," & p.Z & "],")

        sb.AppendLine("""axis"": [" &
            axisVec(0) & "," & axisVec(1) & "," & axisVec(2) & "],")

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

MessageBox.Show("✅ Hole data exported successfully:" & vbLf & outPath)