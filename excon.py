Imports System.IO
Imports System.Text
Imports System.Web.Script.Serialization

Dim asm As AssemblyDocument = ThisDoc.Document
Dim asmDef As AssemblyComponentDefinition = asm.ComponentDefinition

Dim result As New Dictionary(Of String, Object)

' ============================
' OCCURRENCES
' ============================
Dim occList As New List(Of Object)

For Each occ As ComponentOccurrence In asmDef.Occurrences.AllReferencedOccurrences
    Dim o As New Dictionary(Of String, Object)

    o("name") = occ.Name
    o("definition") = occ.Definition.Document.DisplayName
    o("full_path") = occ.Definition.Document.FullFileName
    o("suppressed") = occ.Suppressed
    o("grounded") = occ.Grounded

    ' Transform
    Dim m = occ.Transformation
    Dim mat(3,3) As Double
    For r = 1 To 4
        For c = 1 To 4
            mat(r-1,c-1) = m.Cell(r,c)
        Next
    Next
    o("transform") = mat

    ' Pattern
    If occ.PatternElement IsNot Nothing Then
        o("pattern_parent") = occ.PatternElement.Parent.Name
    Else
        o("pattern_parent") = Nothing
    End If

    occList.Add(o)
Next

result("occurrences") = occList

' ============================
' CONSTRAINTS
' ============================
Dim conList As New List(Of Object)

For Each c As AssemblyConstraint In asmDef.Constraints
    Dim cd As New Dictionary(Of String, Object)

    cd("name") = c.Name
    cd("type") = c.Type
    cd("suppressed") = c.Suppressed

    If TypeOf c Is MateConstraint Or TypeOf c Is FlushConstraint Or TypeOf c Is InsertConstraint Then
        cd("occurrence_1") = c.OccurrenceOne.Name
        cd("occurrence_2") = c.OccurrenceTwo.Name

        cd("entity_1_type") = c.EntityOne.Type
        cd("entity_2_type") = c.EntityTwo.Type

        cd("entity_1_ref") = c.EntityOne.ReferenceKey
        cd("entity_2_ref") = c.EntityTwo.ReferenceKey
    End If

    conList.Add(cd)
Next

result("constraints") = conList

' ============================
' HOLES (PART LEVEL)
' ============================
Dim holeList As New List(Of Object)

For Each occ As ComponentOccurrence In asmDef.Occurrences.AllReferencedOccurrences
    If occ.DefinitionDocumentType <> DocumentTypeEnum.kPartDocumentObject Then Continue For

    Dim part As PartDocument = occ.Definition.Document
    Dim cd = part.ComponentDefinition

    For Each h As HoleFeature In cd.Features.HoleFeatures
        If h.Suppressed Then Continue For

        Dim hd As New Dictionary(Of String, Object)
        hd("occurrence") = occ.Name
        hd("part") = part.DisplayName
        hd("diameter_mm") = h.HoleDiameter.Value * 10
        hd("threaded") = h.Tapped

        Dim axis = h.Axis
        hd("center") = New Double() {axis.RootPoint.X*10, axis.RootPoint.Y*10, axis.RootPoint.Z*10}
        hd("direction") = New Double() {axis.Direction.X, axis.Direction.Y, axis.Direction.Z}

        holeList.Add(hd)
    Next
Next

result("holes") = holeList

' ============================
' EXPORT JSON
' ============================
Dim serializer As New JavaScriptSerializer
Dim json As String = serializer.Serialize(result)

Dim outPath As String = Path.Combine(Path.GetDirectoryName(asm.FullFileName), "assembly_export.json")
File.WriteAllText(outPath, json, Encoding.UTF8)

MessageBox.Show("✅ Assembly exported:" & vbCrLf & outPath, "iLogic Export")