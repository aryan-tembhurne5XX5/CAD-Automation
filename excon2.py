Imports System.IO
Imports System.Text
Imports Inventor
Imports System.Globalization
Imports System.Windows.Forms

Sub Main()

    ' ---------------- 1. Validate Environment ----------------
    Dim oDoc As Document = ThisDoc.Document

    If oDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("Run this rule inside a Part (.ipt) file only.", "Export Error")
        Exit Sub
    End If

    Dim oPartDoc As PartDocument = CType(oDoc, PartDocument)
    Dim oDef As PartComponentDefinition = oPartDoc.ComponentDefinition

    If String.IsNullOrEmpty(oPartDoc.FullFileName) Then
        MessageBox.Show("Please save the IPT file before exporting.", "Export Error")
        Exit Sub
    End If

    ' ---------------- 2. JSON Builder ----------------
    Dim sb As New StringBuilder()
    sb.AppendLine("{")

    ' ---------------- 3. File Info ----------------
    Dim partName As String = Path.GetFileNameWithoutExtension(oPartDoc.FullFileName)

    sb.AppendLine("  ""fileName"": """ & EscapeJson(partName) & """,")
    sb.AppendLine("  ""fullPath"": """ & EscapeJson(oPartDoc.FullFileName) & """,")

    ' ---------------- 4. Physical Properties ----------------
    Try
        Dim mp As MassProperties = oDef.MassProperties

        sb.AppendLine("  ""physical"": {")
        sb.AppendLine("    ""mass_kg"": " & mp.Mass.ToString(CultureInfo.InvariantCulture) & ",")
        sb.AppendLine("    ""volume_cm3"": " & mp.Volume.ToString(CultureInfo.InvariantCulture) & ",")
        sb.AppendLine("    ""area_cm2"": " & mp.Area.ToString(CultureInfo.InvariantCulture))
        sb.AppendLine("  },")
    Catch
        sb.AppendLine("  ""physical"": null,")
    End Try

    ' ---------------- 5. iProperties ----------------
    sb.AppendLine("  ""properties"": {")
    sb.AppendLine("    ""partNumber"": """ & GetProp(oDoc, "Design Tracking Properties", "Part Number") & """,")
    sb.AppendLine("    ""description"": """ & GetProp(oDoc, "Design Tracking Properties", "Description") & """,")
    sb.AppendLine("    ""revision"": """ & GetProp(oDoc, "Summary Information", "Revision Number") & """,")
    sb.AppendLine("    ""title"": """ & GetProp(oDoc, "Summary Information", "Title") & """,")
    sb.AppendLine("    ""material"": """ & GetProp(oDoc, "Design Tracking Properties", "Material") & """")
    sb.AppendLine("  },")

    ' ---------------- 6. Feature List ----------------
    sb.AppendLine("  ""featureList"": [")

    Dim featureNames As New List(Of String)
    For Each f As PartFeature In oDef.Features
        featureNames.Add("""" & EscapeJson(f.Name) & """")
    Next

    sb.AppendLine("    " & String.Join(",", featureNames))
    sb.AppendLine("  ],")

    ' ---------------- 7. Hole / Connection Data ----------------
    sb.AppendLine("  ""connectionPoints"": [")

    Dim connections As New List(Of String)

    For Each oHole As HoleFeature In oDef.Features.HoleFeatures

        If oHole.Suppressed Then Continue For

        Dim placement As HolePlacementDefinition = oHole.PlacementDefinition

        If placement.Type <> ObjectTypeEnum.kSketchHolePlacementDefinitionObject Then Continue For

        Dim sketchPlacement As SketchHolePlacementDefinition =
            CType(placement, SketchHolePlacementDefinition)

        Dim isThreaded As Boolean = oHole.Tapped
        Dim holeType As String = If(isThreaded, "Threaded", "Simple")

        Dim diameterMm As Double = 0
        Try
            diameterMm = oHole.HoleDiameter.Value * 10.0 ' cm → mm
        Catch
        End Try

        For Each sp As SketchPoint In sketchPlacement.SketchCenterPoints

            Dim pt As Point = sp.Geometry3d

            Dim px As Double = pt.X * 10.0
            Dim py As Double = pt.Y * 10.0
            Dim pz As Double = pt.Z * 10.0

            Dim ps As PlanarSketch = CType(sp.Parent, PlanarSketch)
            Dim plane As Plane = ps.PlanarEntityGeometry

            Dim cb As New StringBuilder()
            cb.AppendLine("    {")
            cb.AppendLine("      ""featureName"": """ & EscapeJson(oHole.Name) & """,")
            cb.AppendLine("      ""type"": """ & holeType & """,")
            cb.AppendLine("      ""diameterMm"": " & diameterMm.ToString(CultureInfo.InvariantCulture) & ",")
            cb.AppendLine("      ""threaded"": " & isThreaded.ToString().ToLower() & ",")
            cb.AppendLine("      ""center"": { ""x"": " &
                          px.ToString(CultureInfo.InvariantCulture) & ", ""y"": " &
                          py.ToString(CultureInfo.InvariantCulture) & ", ""z"": " &
                          pz.ToString(CultureInfo.InvariantCulture) & " },")
            cb.AppendLine("      ""axis"": { ""x"": " &
                          plane.Normal.X.ToString(CultureInfo.InvariantCulture) & ", ""y"": " &
                          plane.Normal.Y.ToString(CultureInfo.InvariantCulture) & ", ""z"": " &
                          plane.Normal.Z.ToString(CultureInfo.InvariantCulture) & " }")
            cb.Append("    }")

            connections.Add(cb.ToString())
        Next
    Next

    sb.AppendLine(String.Join("," & vbCrLf, connections))
    sb.AppendLine("  ]")
    sb.AppendLine("}")

    ' ---------------- 8. Write File ----------------
    Dim jsonPath As String = Path.ChangeExtension(oPartDoc.FullFileName, "json")

    Try
        File.WriteAllText(jsonPath, sb.ToString())
        MessageBox.Show("Export successful!" & vbCrLf & jsonPath, "iLogic Export")
    Catch ex As Exception
        MessageBox.Show("File write error: " & ex.Message, "Export Error")
    End Try

End Sub

' ---------------- Helper Functions ----------------

Function GetProp(doc As Document, setName As String, propName As String) As String
    Try
        Return EscapeJson(doc.PropertySets.Item(setName).Item(propName).Value.ToString())
    Catch
        Return ""
    End Try
End Function

Function EscapeJson(value As String) As String
    If String.IsNullOrEmpty(value) Then Return ""
    Return value.Replace("\", "\\").Replace("""", "\""").
                 Replace(vbCrLf, "\n").Replace(vbCr, "\n").Replace(vbLf, "\n")
End Function