' ================================
' SAFE iLogic Assembly Export
' ================================

Imports System.Text
Imports System.IO

Dim asmDoc As AssemblyDocument = ThisApplication.ActiveDocument
Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

Dim sb As New StringBuilder()
sb.AppendLine("{")
sb.AppendLine("""assembly"": """ & asmDoc.DisplayName & """,")

' ================================
' OCCURRENCES
' ================================
sb.AppendLine("""occurrences"": [")

Dim firstOcc As Boolean = True

For Each occ As ComponentOccurrence In asmDef.Occurrences
    If Not firstOcc Then sb.AppendLine(",")
    firstOcc = False

    Dim m = occ.Transformation

    sb.AppendLine("{")
    sb.AppendLine("""name"": """ & occ.Name & """,")
    sb.AppendLine("""definition"": """ & occ.Definition.Document.DisplayName & """,")
    sb.AppendLine("""suppressed"": " & occ.Suppressed.ToString().ToLower() & ",")
    sb.AppendLine("""grounded"": " & occ.Grounded.ToString().ToLower() & ",")

    sb.AppendLine("""transform"": [")
    For r = 1 To 4
        sb.Append("  [")
        For c = 1 To 4
            sb.Append(m.Cell(r, c))
            If c < 4 Then sb.Append(", ")
        Next
        sb.Append("]")
        If r < 4 Then sb.AppendLine(",") Else sb.AppendLine()
    Next
    sb.AppendLine("]")

    sb.Append("}")
Next

sb.AppendLine("],")

' ================================
' CONSTRAINTS
' ================================
sb.AppendLine("""constraints"": [")

Dim firstC As Boolean = True

For Each c As AssemblyConstraint In asmDef.Constraints
    If Not firstC Then sb.AppendLine(",")
    firstC = False

    sb.AppendLine("{")
    sb.AppendLine("""name"": """ & c.Name & """,")
    sb.AppendLine("""type"": " & c.Type & ",")
    sb.AppendLine("""suppressed"": " & c.Suppressed.ToString().ToLower())
    sb.Append("}")
Next

sb.AppendLine("]")

sb.AppendLine("}")

' ================================
' WRITE FILE
' ================================
Dim outPath As String =
    System.IO.Path.Combine(
        System.IO.Path.GetDirectoryName(asmDoc.FullFileName),
        "assembly_export.json"
    )

System.IO.File.WriteAllText(outPath, sb.ToString())

MessageBox.Show("✅ Assembly exported to:" & vbLf & outPath)