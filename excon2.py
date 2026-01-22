' ==========================================
' PART DATA EXPORT — ASSEMBLY READY
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
' ORIGIN & AXES
' -------------------------
Dim wp = cd.WorkPoints.Item("Center Point")
sb.AppendLine("""origin"": [" &
    wp.Point.X & "," &
    wp.Point.Y & "," &
    wp.Point.Z & "],")

sb.AppendLine("""axes"": {")
For Each ax As WorkAxis In cd.WorkAxes
    sb.AppendLine("""" & ax.Name & """: {")
    sb.AppendLine("""origin"": [" &
        ax.Line.RootPoint.X & "," &
        ax.Line.RootPoint.Y & "," &
        ax.Line.RootPoint.Z & "],")
    sb.AppendLine("""direction"": [" &
        ax.Line.Direction.X & "," &
        ax.Line.Direction.Y & "," &
        ax.Line.Direction.Z & "]")
    sb.AppendLine("},")
Next
sb.Remove(sb.Length - 3, 1)
sb.AppendLine("},")

' -------------------------
' HOLES (CRITICAL)
' -------------------------
sb.AppendLine("""holes"": [")

Dim firstHole As Boolean = True

For Each h As HoleFeature In cd.Features.HoleFeatures
    If h.Suppressed Then Continue For

    If Not firstHole Then sb.AppendLine(",")
    firstHole = False

    Dim axis = h.Axis
    Dim hdef = h.HoleDefinition

    sb.AppendLine("{")
    sb.AppendLine("""name"": """ & h.Name & """,")

    ' Diameter (cm → mm)
    Dim diaMM As Double = 0
    Try
        diaMM = hdef.Diameter.Value * 10
    Catch
    End Try
    sb.AppendLine("""diameter_mm"": " & diaMM & ",")

    ' Hole center (3D)
    sb.AppendLine("""center"": [" &
        axis.RootPoint.X & "," &
        axis.RootPoint.Y & "," &
        axis.RootPoint.Z & "],")

    ' Hole axis
    sb.AppendLine("""axis"": [" &
        axis.Direction.X & "," &
        axis.Direction.Y & "," &
        axis.Direction.Z & "],")

    sb.AppendLine("""threaded"": " & h.Tapped.ToString().ToLower() & ",")

    sb.AppendLine("""extent"": " & hdef.ExtentType)

    sb.Append("}")
Next

sb.AppendLine("],")

' -------------------------
' BOUNDING BOX
' -------------------------
Dim box = cd.RangeBox
sb.AppendLine("""bounding_box"": {")
sb.AppendLine("""min"": [" &
    box.MinPoint.X & "," &
    box.MinPoint.Y & "," &
    box.MinPoint.Z & "],")
sb.AppendLine("""max"": [" &
    box.MaxPoint.X & "," &
    box.MaxPoint.Y & "," &
    box.MaxPoint.Z & "]")
sb.AppendLine("}")

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

MessageBox.Show("✅ Part exported for assembly:" & vbLf & outPath)