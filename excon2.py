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

    ' --- 3. General Info & Physical Properties ---
    Dim partName As String = System.IO.Path.GetFileNameWithoutExtension(oPartDoc.FullFileName)
    sb.AppendLine("  ""fileName"": """ & EscapeJson(partName) & """,")
    sb.AppendLine("  ""fullPath"": """ & EscapeJson(oPartDoc.FullFileName) & """,")
    
    ' Physical (MassProps)
    Try
        Dim mass As Double = oPartDoc.ComponentDefinition.MassProperties.Mass ' kg
        Dim volume As Double = oPartDoc.ComponentDefinition.MassProperties.Volume ' cm^3
        Dim area As Double = oPartDoc.ComponentDefinition.MassProperties.Area ' cm^2
        
        sb.AppendLine("  ""physical"": {")
        sb.AppendLine("    ""mass_kg"": " & mass.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
        sb.AppendLine("    ""volume_cm3"": " & volume.ToString(System.Globalization.CultureInfo.InvariantCulture) & ",")
        sb.AppendLine("    ""area_cm2"": " & area.ToString(System.Globalization.CultureInfo.InvariantCulture))
        sb.AppendLine("  },")
    Catch
        sb.AppendLine("  ""physical"": null,")
    End Try

    ' --- 4. iProperties ---
    sb.AppendLine("  ""properties"": {")
    sb.AppendLine("    ""partNumber"": """ & GetProp(oDoc, "Design Tracking Properties", "Part Number") & """,")
    sb.AppendLine("    ""description"": """ & GetProp(oDoc, "Design Tracking Properties", "Description") & """,")
    sb.AppendLine("    ""revision"": """ & GetProp(oDoc, "Summary Information", "Revision Number") & """,")
    sb.AppendLine("    ""title"": """ & GetProp(oDoc, "Summary Information", "Title") & """,")
    sb.AppendLine("    ""material"": """ & GetProp(oDoc, "Design Tracking Properties", "Material") & """")
    sb.AppendLine("  },")

    ' --- 5. Feature List (Overview) ---
    sb.AppendLine("  ""featureList"": [")
    Dim featList As New List(Of String)
    For Each feat As PartFeature In oDef.Features
        featList.Add("""" & EscapeJson(feat.Name) & """")
    Next
    sb.Append(String.Join(",", featList.ToArray()))
    sb.AppendLine("  ],")

    ' --- 6. DETAILED HOLE DATA ---
    sb.AppendLine("  ""connectionPoints"": [")
    
    Dim connectionList As New List(Of String)

    ' Iterate through standard HoleFeatures
    For Each oHole As HoleFeature In oDef.Features.HoleFeatures
        If oHole.Suppressed Then Continue For

        ' FIX: Correctly check if the placement definition is a Sketch Hole using ObjectTypeEnum
        If oHole.PlacementDefinition.Type = ObjectTypeEnum.kSketchHolePlacementDefinitionObject Then
            
            Dim isThreaded As Boolean = oHole.Tapped
            Dim holeType As String = If(isThreaded, "Threaded", "Simple")
            Dim diameterMm As Double = 0.0

            Try
                diameterMm = oHole.HoleDiameter.Value * 10.0 ' cm to mm
            Catch
                diameterMm = 0.0
            End Try

            Dim oSketchPoints As Object = oHole.SketchCenterPoints

            For Each oPoint As SketchPoint In oSketchPoints
                
                ' 3D Center Calculation
                Dim pX As Double = oPoint.Geometry3d.X * 10.0
                Dim pY As Double = oPoint.Geometry3d.Y * 10.0
                Dim pZ As Double = oPoint.Geometry3d.Z * 10.0

                ' Axis Vector Calculation
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
                cb.AppendLine("      ""threaded"": " & isThreaded.ToString().ToLower() & ",")
                
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

    ' --- 7. File Output ---
    Dim jsonPath As String = System.IO.Path.ChangeExtension(oPartDoc.FullFileName, "json")
    
    If String.IsNullOrEmpty(oPartDoc.FullFileName) Then
        MessageBox.Show("Please save the IPT file before exporting.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Exit Sub
    End If

    Try
        System.IO.File.WriteAllText(jsonPath, sb.ToString())
        MessageBox.Show("Data exported successfully!" & vbCrLf & jsonPath, "iLogic Export")
    Catch ex As Exception
        MessageBox.Show("Error writing file: " & ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

End Sub

' --- Helper Functions ---

Function GetProp(doc As Document, setName As String, propName As String) As String
    Try
        Return EscapeJson(doc.PropertySets.Item(setName).Item(propName).Value.ToString())
    Catch
        Return ""
    End Try
End Function

Function EscapeJson(str As String) As String
    If str Is Nothing Then Return ""
    Return str.Replace("\", "\\").Replace("""", "\""").Replace(vbCrLf, "\n").Replace(vbCr, "\n").Replace(vbLf, "\n")
End Function