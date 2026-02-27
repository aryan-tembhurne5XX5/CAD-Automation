using Inventor;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
// NOTE: Do NOT add "using System.IO" — Inventor.File / Inventor.Path conflict.
//       Use System.IO.File, System.IO.Path, System.IO.Directory explicitly.

namespace InventorAssemblyReconstructor
{
    class Program
    {
        static Inventor.Application invApp;
        static TransientGeometry tg;

        static void Main(string[] args)
        {
            string jsonPath = @"E:\Phase 1\Assembly 2\assembly2_full_export.json";
            string outputPath = @"E:\Phase 1\Assembly 2\Reconstructed\Reconstructed.iam";

            // ── Connect to Inventor ──────────────────────────────────────────────
            try
            {
                invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
                Console.WriteLine("Connected to running Inventor instance.");
            }
            catch
            {
                Type t = Type.GetTypeFromProgID("Inventor.Application");
                invApp = (Inventor.Application)Activator.CreateInstance(t);
                invApp.Visible = true;
                Console.WriteLine("Started new Inventor instance.");
            }

            tg = invApp.TransientGeometry;

            // ── Read JSON ───────────────────────────────────────────────────────
            string json = System.IO.File.ReadAllText(jsonPath);
            AssemblyExport export = JsonConvert.DeserializeObject<AssemblyExport>(json);

            Console.WriteLine($"Loaded JSON: {export.assembly_metadata.assembly_name}");
            Console.WriteLine($"  Components : {export.components.Count}");
            Console.WriteLine($"  Constraints: {export.constraints.Count}");

            // ── Create new assembly document ─────────────────────────────────────
            string outputDir = System.IO.Path.GetDirectoryName(outputPath);
            if (!System.IO.Directory.Exists(outputDir)) System.IO.Directory.CreateDirectory(outputDir);

            AssemblyDocument asmDoc = (AssemblyDocument)invApp.Documents.Add(
                DocumentTypeEnum.kAssemblyDocumentObject,
                invApp.FileManager.GetTemplateFile(DocumentTypeEnum.kAssemblyDocumentObject),
                true);

            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;

            // ── Place occurrences ────────────────────────────────────────────────
            // occurrenceMap: occurrence_name → ComponentOccurrence (for constraint binding)
            var occurrenceMap = new Dictionary<string, ComponentOccurrence>(StringComparer.OrdinalIgnoreCase);

            Console.WriteLine("\nPlacing components...");
            PlaceOccurrences(export.components, asmDef, occurrenceMap, "");

            // ── Apply constraints ────────────────────────────────────────────────
            Console.WriteLine("\nApplying constraints...");
            int constraintOk = 0;
            int constraintFail = 0;

            foreach (ConstraintData con in export.constraints)
            {
                try
                {
                    ApplyConstraint(con, asmDef, occurrenceMap);
                    constraintOk++;
                    Console.WriteLine($"  [OK]   {con.constraint_name} ({con.constraint_type})");
                }
                catch (Exception ex)
                {
                    constraintFail++;
                    Console.WriteLine($"  [FAIL] {con.constraint_name} ({con.constraint_type}): {ex.Message}");
                }
            }

            // ── Save ─────────────────────────────────────────────────────────────
            asmDoc.SaveAs(outputPath, false);
            Console.WriteLine($"\nSaved → {outputPath}");
            Console.WriteLine($"Constraints: {constraintOk} OK, {constraintFail} failed");
        }

        // ══════════════════════════════════════════════════════════════════════════
        // PLACE OCCURRENCES (recursive for sub-assemblies)
        // ══════════════════════════════════════════════════════════════════════════

        static void PlaceOccurrences(
            List<ComponentData> components,
            AssemblyComponentDefinition asmDef,
            Dictionary<string, ComponentOccurrence> map,
            string parentPath)
        {
            foreach (ComponentData comp in components)
            {
                if (comp.suppressed) continue;
                if (!System.IO.File.Exists(comp.file_name))
                {
                    Console.WriteLine($"  [SKIP] File not found: {comp.file_name}");
                    continue;
                }

                // Build placement matrix from stored transform
                Matrix placeMx = BuildMatrix(comp.transform);

                ComponentOccurrence occ = asmDef.Occurrences.Add(comp.file_name, placeMx);

                // Track by occurrence name and by full path
                string key = parentPath == "" ? comp.occurrence_name : parentPath + "/" + comp.occurrence_name;
                if (!map.ContainsKey(comp.occurrence_name)) map[comp.occurrence_name] = occ;
                if (!map.ContainsKey(key)) map[key] = occ;

                Console.WriteLine($"  Placed: {comp.occurrence_name}");

                // Grounded?
                if (comp.grounded) occ.Grounded = true;

                // Sub-assemblies — place their children too (they will be placed
                // inside the sub-assembly document, not in this one, so we only
                // recurse to register names for constraint lookup)
                if (comp.sub_components != null && comp.sub_components.Count > 0)
                    RegisterSubOccurrences(comp.sub_components, occ, map, key);
            }
        }

        // For sub-assembly children we can't place them (they are inside the sub-asm
        // document), but we do register their proxies so constraints can resolve them.
        static void RegisterSubOccurrences(
            List<ComponentData> subs,
            ComponentOccurrence parentOcc,
            Dictionary<string, ComponentOccurrence> map,
            string parentPath)
        {
            foreach (ComponentData sub in subs)
            {
                string key = parentPath + "/" + sub.occurrence_name;
                // The sub-occurrence exists inside the parent's definition.
                // We locate it through the parent's sub-occurrences.
                try
                {
                    foreach (ComponentOccurrence subOcc in parentOcc.SubOccurrences)
                    {
                        if (subOcc.Name == sub.occurrence_name || subOcc.Name.StartsWith(sub.occurrence_name + ":"))
                        {
                            if (!map.ContainsKey(sub.occurrence_name)) map[sub.occurrence_name] = subOcc;
                            if (!map.ContainsKey(key)) map[key] = subOcc;
                            break;
                        }
                    }
                }
                catch { }

                if (sub.sub_components != null && sub.sub_components.Count > 0)
                    RegisterSubOccurrences(sub.sub_components, parentOcc, map, key);
            }
        }

        // ══════════════════════════════════════════════════════════════════════════
        // BUILD MATRIX from stored rotation + translation
        // ══════════════════════════════════════════════════════════════════════════

        static Matrix BuildMatrix(TransformData t)
        {
            Matrix m = tg.CreateMatrix();

            if (t == null) return m; // identity

            // Rotation (3×3 sub-matrix)
            if (t.rotation_matrix != null)
            {
                for (int i = 0; i < 3; i++)
                    for (int j = 0; j < 3; j++)
                        m.set_Cell(i + 1, j + 1, t.rotation_matrix[i][j]);
            }

            // Translation (column 4)
            if (t.translation_cm != null)
            {
                m.set_Cell(1, 4, t.translation_cm[0]);
                m.set_Cell(2, 4, t.translation_cm[1]);
                m.set_Cell(3, 4, t.translation_cm[2]);
            }

            m.set_Cell(4, 4, 1.0); // homogeneous

            return m;
        }

        // ══════════════════════════════════════════════════════════════════════════
        // APPLY CONSTRAINT
        // Strategy:
        //   1. Try reference-key binding (most accurate — works when keys are valid).
        //   2. Fall back to geometric-signature matching (area/length/normal).
        //   3. Fall back to first-face heuristic so something is placed.
        // ══════════════════════════════════════════════════════════════════════════

        static void ApplyConstraint(
            ConstraintData con,
            AssemblyComponentDefinition asmDef,
            Dictionary<string, ComponentOccurrence> map)
        {
            if (con.suppressed) return;

            object entity1 = ResolveEntity(con.entity_one, con.occurrence_one, map, asmDef);
            object entity2 = ResolveEntity(con.entity_two, con.occurrence_two, map, asmDef);

            if (entity1 == null) throw new Exception("Could not resolve entity_one");
            if (entity2 == null) throw new Exception("Could not resolve entity_two");

            // constraint_type strings from extractor: c.Type.ToString() → e.g. "kMateConstraintObject"
            switch (con.constraint_type)
            {
                // ── Mate ──────────────────────────────────────────────────────────
                // REAL signature (Autodesk docs + community code):
                //   AddMateConstraint(EntityOne, EntityTwo, Offset,
                //       [EntityOneInferredType As InferredTypeEnum],
                //       [EntityTwoInferredType As InferredTypeEnum],
                //       [BiasPointOne], [BiasPointTwo])
                // No MateConstraintType/MateConstraintTypeEnum parameter exists.
                case "kMateConstraintObject":
                case "kMateConstraint":
                    {
                        double offset = con.offset_cm ?? 0.0;
                        asmDef.Constraints.AddMateConstraint(
                            entity1, entity2, offset,
                            InferredTypeEnum.kNoInference,
                            InferredTypeEnum.kNoInference,
                            Type.Missing,
                            Type.Missing);
                        break;
                    }

                // ── Flush ─────────────────────────────────────────────────────────
                // REAL signature: AddFlushConstraint(EntityOne, EntityTwo, Offset)
                case "kFlushConstraintObject":
                case "kFlushConstraint":
                    {
                        double offset = con.offset_cm ?? 0.0;
                        asmDef.Constraints.AddFlushConstraint(entity1, entity2, offset);
                        break;
                    }

                // ── Angle ─────────────────────────────────────────────────────────
                // REAL signature (verified from Autodesk community working code):
                //   AddAngleConstraint(EntityOne, EntityTwo, Angle)
                // SolutionType is an optional 4th argument. Every working community
                // example omits it and uses only 3 args. The enum member names are
                // not publicly documented and guessing them reliably is impossible,
                // so we use the 3-argument form which defaults to "directed" mode.
                case "kAngleConstraintObject":
                case "kAngleConstraint":
                    {
                        double angle = con.angle_rad ?? 0.0;
                        asmDef.Constraints.AddAngleConstraint(entity1, entity2, angle);
                        break;
                    }

                // ── Tangent ───────────────────────────────────────────────────────
                // REAL signature:
                //   AddTangentConstraint(EntityOne, EntityTwo,
                //       InsideTangency As Boolean, InsideSecondEntity As Boolean)
                case "kTangentConstraintObject":
                case "kTangentConstraint":
                    {
                        bool inside = con.inside_tangency ?? false;
                        asmDef.Constraints.AddTangentConstraint(entity1, entity2, inside, false);
                        break;
                    }

                // ── Insert ────────────────────────────────────────────────────────
                // REAL signature (verified from iLogicRules community source):
                //   AddInsertConstraint2(EntityOne, EntityTwo,
                //       AxesOpposed As Boolean,
                //       Distance,
                //       LockRotation As Boolean)
                // Note: AddInsertConstraint (v1) has different arg order — use v2.
                case "kInsertConstraintObject":
                case "kInsertConstraint":
                    {
                        double offset = con.offset_cm ?? 0.0;
                        asmDef.Constraints.AddInsertConstraint2(
                            entity1, entity2,
                            false,   // AxesOpposed — false = aligned, true = opposed
                            offset,
                            false);  // LockRotation
                        break;
                    }

                // ── Symmetry ──────────────────────────────────────────────────────
                case "kSymmetryConstraintObject":
                case "kSymmetryConstraint":
                    throw new NotSupportedException(
                        "SymmetryConstraint needs a 3rd plane entity not stored in JSON — skipped");

                default:
                    throw new NotSupportedException($"Unknown constraint type: {con.constraint_type}");
            }
        }

        // ══════════════════════════════════════════════════════════════════════════
        // RESOLVE ENTITY
        //   Priority:
        //     1. Reference-key bind (exact, using stored reference_key_string)
        //     2. Geometric signature match (area for faces, length for edges)
        //     3. First available face/edge (last resort)
        // ══════════════════════════════════════════════════════════════════════════

        static object ResolveEntity(
            EntityData ed,
            string occurrenceName,
            Dictionary<string, ComponentOccurrence> map,
            AssemblyComponentDefinition asmDef)
        {
            if (ed == null) return null;

            // Find the owning occurrence
            ComponentOccurrence occ = null;
            if (occurrenceName != null) map.TryGetValue(occurrenceName, out occ);

            // ── 1. Reference-key binding ──────────────────────────────────────────
            if (!string.IsNullOrEmpty(ed.reference_key_string) &&
                !string.IsNullOrEmpty(ed.context_key_string) &&
                !string.IsNullOrEmpty(ed.owner_document))
            {
                try
                {
                    // Open/find the owner part document
                    Document ownerDoc = FindOrOpenDocument(ed.owner_document);
                    if (ownerDoc != null)
                    {
                        ReferenceKeyManager mgr = ownerDoc.ReferenceKeyManager;

                        byte[] refKey = new byte[1];
                        mgr.StringToKey(ed.reference_key_string, ref refKey);

                        byte[] ctxKey = new byte[1];
                        mgr.StringToKey(ed.context_key_string, ref ctxKey);

                        int restoredCtx = mgr.LoadContextFromArray(ref ctxKey);

                        object matchType;
                        object nativeEntity = mgr.BindKeyToObject(refKey, restoredCtx, out matchType);

                        if (nativeEntity != null && occ != null)
                        {
                            // Create proxy relative to the assembly occurrence
                            object proxy = CreateProxy(nativeEntity, occ, asmDef);
                            if (proxy != null) return proxy;
                        }
                    }
                }
                catch { /* fall through to geometric match */ }
            }

            // ── 2. Geometric-signature match ──────────────────────────────────────
            if (occ != null)
            {
                try
                {
                    object geoMatch = GeometricMatch(ed, occ);
                    if (geoMatch != null) return geoMatch;
                }
                catch { }
            }

            // ── 3. Last resort: first face or edge of the occurrence ───────────────
            if (occ != null)
            {
                try
                {
                    SurfaceBodies bodies = occ.SurfaceBodies;
                    if (bodies.Count > 0)
                    {
                        SurfaceBody body = bodies[1];

                        if (ed.entity_type != null && ed.entity_type.Contains("Edge") && body.Edges.Count > 0)
                            return body.Edges[1];

                        if (body.Faces.Count > 0)
                            return body.Faces[1];
                    }
                }
                catch { }
            }

            return null;
        }

        // ══════════════════════════════════════════════════════════════════════════
        // GEOMETRIC MATCH
        //   For faces  : match surface_type + area (within tolerance)
        //   For edges  : match curve_type  + length (within tolerance)
        //   Normal / midpoint used as tie-breaker when multiple candidates match.
        // ══════════════════════════════════════════════════════════════════════════

        static object GeometricMatch(EntityData ed, ComponentOccurrence occ)
        {
            bool isFace = ed.entity_type != null &&
                          (ed.entity_type == "FaceProxy" || ed.entity_type == "Face");
            bool isEdge = ed.entity_type != null &&
                          (ed.entity_type == "EdgeProxy" || ed.entity_type == "Edge");

            SurfaceBodies bodies = occ.SurfaceBodies;

            if (isFace && ed.area_cm2.HasValue)
            {
                double targetArea = ed.area_cm2.Value;
                double bestDelta = double.MaxValue;
                Face bestFace = null;

                foreach (SurfaceBody body in bodies)
                {
                    foreach (Face face in body.Faces)
                    {
                        // Filter by surface type if available
                        if (!string.IsNullOrEmpty(ed.surface_type) &&
                            face.SurfaceType.ToString() != ed.surface_type)
                            continue;

                        double delta = Math.Abs(face.Evaluator.Area - targetArea);
                        if (delta < bestDelta)
                        {
                            // Optionally refine with normal direction match
                            if (ed.face_normal_at_center != null)
                            {
                                try
                                {
                                    SurfaceEvaluator eval = face.Evaluator;
                                    Box2d uvRect = eval.ParamRangeRect;
                                    double[] pars = new double[2]
                                    {
                                        (uvRect.MinPoint.X + uvRect.MaxPoint.X) / 2.0,
                                        (uvRect.MinPoint.Y + uvRect.MaxPoint.Y) / 2.0
                                    };
                                    double[] normals = new double[3];
                                    eval.GetNormal(ref pars, ref normals);

                                    double dot = normals[0] * ed.face_normal_at_center[0]
                                               + normals[1] * ed.face_normal_at_center[1]
                                               + normals[2] * ed.face_normal_at_center[2];

                                    // Weight: combine area proximity and normal alignment
                                    double weightedDelta = delta - dot * 0.001;
                                    if (weightedDelta < bestDelta)
                                    {
                                        bestDelta = weightedDelta;
                                        bestFace = face;
                                    }
                                    continue;
                                }
                                catch { }
                            }

                            bestDelta = delta;
                            bestFace = face;
                        }
                    }
                }

                if (bestFace != null)
                {
                    // Return proxy if possible
                    try
                    {
                        object proxy;
                        occ.CreateGeometryProxy(bestFace, out proxy);
                        return proxy;
                    }
                    catch { return bestFace; }
                }
            }

            if (isEdge && ed.length_cm.HasValue)
            {
                double targetLen = ed.length_cm.Value;
                double bestDelta = double.MaxValue;
                Edge bestEdge = null;

                foreach (SurfaceBody body in bodies)
                {
                    foreach (Edge edge in body.Edges)
                    {
                        if (!string.IsNullOrEmpty(ed.curve_type) &&
                            edge.CurveType.ToString() != ed.curve_type)
                            continue;

                        double len = GetEdgeLength(edge);
                        double delta = Math.Abs(len - targetLen);

                        if (delta < bestDelta)
                        {
                            // Refine with midpoint proximity
                            if (ed.edge_midpoint != null)
                            {
                                try
                                {
                                    CurveEvaluator eval = edge.Evaluator;
                                    double startP, endP;
                                    eval.GetParamExtents(out startP, out endP);
                                    double mid = (startP + endP) / 2.0;
                                    double[] pArr = { mid };
                                    double[] midPt = new double[3];
                                    eval.GetPointAtParam(ref pArr, ref midPt);

                                    double dx = midPt[0] - ed.edge_midpoint[0];
                                    double dy = midPt[1] - ed.edge_midpoint[1];
                                    double dz = midPt[2] - ed.edge_midpoint[2];
                                    double ptDist = Math.Sqrt(dx * dx + dy * dy + dz * dz);

                                    double weightedDelta = delta + ptDist * 0.01;
                                    if (weightedDelta < bestDelta)
                                    {
                                        bestDelta = weightedDelta;
                                        bestEdge = edge;
                                    }
                                    continue;
                                }
                                catch { }
                            }

                            bestDelta = delta;
                            bestEdge = edge;
                        }
                    }
                }

                if (bestEdge != null)
                {
                    try
                    {
                        object proxy;
                        occ.CreateGeometryProxy(bestEdge, out proxy);
                        return proxy;
                    }
                    catch { return bestEdge; }
                }
            }

            return null;
        }

        // ══════════════════════════════════════════════════════════════════════════
        // CREATE PROXY from native entity + occurrence
        // ══════════════════════════════════════════════════════════════════════════

        static object CreateProxy(object native, ComponentOccurrence occ, AssemblyComponentDefinition asmDef)
        {
            try
            {
                object proxy;
                occ.CreateGeometryProxy(native, out proxy);
                return proxy;
            }
            catch { }

            // Try work-feature proxies
            try
            {
                if (native is WorkPlane)
                {
                    object proxy;
                    occ.CreateGeometryProxy(native, out proxy);
                    return proxy;
                }
                if (native is WorkAxis)
                {
                    object proxy;
                    occ.CreateGeometryProxy(native, out proxy);
                    return proxy;
                }
            }
            catch { }

            return null;
        }

        // ══════════════════════════════════════════════════════════════════════════
        // FIND OR OPEN DOCUMENT
        // ══════════════════════════════════════════════════════════════════════════

        static Document FindOrOpenDocument(string fullPath)
        {
            // Check if already open
            foreach (Document doc in invApp.Documents)
            {
                if (string.Equals(doc.FullFileName, fullPath, StringComparison.OrdinalIgnoreCase))
                    return doc;
            }

            // Open silently if file exists
            if (System.IO.File.Exists(fullPath))
            {
                try { return invApp.Documents.Open(fullPath, false); }
                catch { }
            }

            return null;
        }

        // ══════════════════════════════════════════════════════════════════════════
        // EDGE LENGTH (mirrors extractor logic)
        // ══════════════════════════════════════════════════════════════════════════

        static double GetEdgeLength(Edge edge)
        {
            try
            {
                CurveEvaluator eval = edge.Evaluator;
                if (edge.CurveType == CurveTypeEnum.kLineCurve)
                {
                    Point p1 = edge.StartVertex.Point;
                    Point p2 = edge.StopVertex.Point;
                    double dx = p2.X - p1.X, dy = p2.Y - p1.Y, dz = p2.Z - p1.Z;
                    return Math.Sqrt(dx * dx + dy * dy + dz * dz);
                }
                else if (edge.CurveType == CurveTypeEnum.kCircleCurve)
                {
                    Circle circ = (Circle)edge.Geometry;
                    double start, end;
                    eval.GetParamExtents(out start, out end);
                    return circ.Radius * Math.Abs(end - start);
                }
                else
                {
                    double start, end;
                    eval.GetParamExtents(out start, out end);
                    return Math.Abs(end - start);
                }
            }
            catch { return 0.0; }
        }
    }

    // ══════════════════════════════════════════════════════════════════════════════
    // DATA STRUCTURES  (must match Program.cs exactly)
    // ══════════════════════════════════════════════════════════════════════════════

    public class AssemblyExport
    {
        public AssemblyMetadata assembly_metadata { get; set; }
        public List<ComponentData> components { get; set; } = new List<ComponentData>();
        public List<ConstraintData> constraints { get; set; } = new List<ConstraintData>();
    }

    public class AssemblyMetadata
    {
        public string assembly_name { get; set; }
        public string full_file_name { get; set; }
        public string internal_name { get; set; }
        public int total_occurrences { get; set; }
        public int total_constraints { get; set; }
    }

    public class ComponentData
    {
        public string occurrence_name { get; set; }
        public string occurrence_path { get; set; }
        public string file_name { get; set; }
        public string component_type { get; set; }
        public bool grounded { get; set; }
        public bool suppressed { get; set; }
        public bool visible { get; set; }
        public TransformData transform { get; set; }
        public List<ComponentData> sub_components { get; set; } = new List<ComponentData>();
    }

    public class TransformData
    {
        public double[][] rotation_matrix { get; set; }
        public double[] translation_cm { get; set; }
    }

    public class ConstraintData
    {
        public string constraint_name { get; set; }
        public string constraint_type { get; set; }
        public string constraint_extra_info { get; set; }
        public bool suppressed { get; set; }
        public string occurrence_one { get; set; }
        public string occurrence_two { get; set; }
        public double? offset_cm { get; set; }
        public double? angle_rad { get; set; }
        public string mate_type { get; set; }
        public string angle_type { get; set; }
        public bool? inside_tangency { get; set; }
        public EntityData entity_one { get; set; }
        public EntityData entity_two { get; set; }
    }

    public class EntityData
    {
        public string entity_type { get; set; }
        public string proxy_context_occurrence { get; set; }
        public string owner_document { get; set; }
        public string reference_key_string { get; set; }
        public string context_key_string { get; set; }
        public bool? bind_succeeded { get; set; }
        public string bind_match_type { get; set; }
        public string key_extraction_error { get; set; }
        public string surface_type { get; set; }
        public double? area_cm2 { get; set; }
        public double[] face_normal_at_center { get; set; }
        public double[] face_point_at_center { get; set; }
        public double[] face_bbox_min { get; set; }
        public double[] face_bbox_max { get; set; }
        public string curve_type { get; set; }
        public double? length_cm { get; set; }
        public double[] edge_midpoint { get; set; }
        public double[] edge_tangent_at_mid { get; set; }
        public string work_feature_name { get; set; }
    }
}