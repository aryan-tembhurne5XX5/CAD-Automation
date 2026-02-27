using Inventor;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace InventorAssemblyExporter
{
    class Program
    {
        static Inventor.Application invApp;

        static void Main(string[] args)
        {
            string phaseRoot = @"E:\Phase 1";

            // ── Connect to Inventor ──────────────────────────────────────────
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

            // Loop Assembly 1 → Assembly 8
            for (int i = 1; i <= 8; i++)
            {
                string assemblyFolder = System.IO.Path.Combine(phaseRoot, $"Assembly {i}");
                string outputRoot = System.IO.Path.Combine(phaseRoot, $"AAs{i}", "extAss");

                if (!Directory.Exists(assemblyFolder))
                {
                    Console.WriteLine($"Skipping missing folder: {assemblyFolder}");
                    continue;
                }

                Directory.CreateDirectory(outputRoot);

                // Get ALL .iam files recursively
                string[] iamFiles = Directory.GetFiles(assemblyFolder, "*.iam", SearchOption.AllDirectories);

                foreach (string assemblyPath in iamFiles)
                {
                    try
                    {
                        Console.WriteLine("Processing: " + assemblyPath);

                        AssemblyDocument asmDoc =
                            (AssemblyDocument)invApp.Documents.Open(assemblyPath, true);

                        AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;

                        AssemblyExport export = new AssemblyExport
                        {
                            assembly_metadata = new AssemblyMetadata
                            {
                                assembly_name = asmDoc.DisplayName,
                                full_file_name = asmDoc.FullFileName,
                                internal_name = asmDoc.InternalName,
                                total_occurrences = asmDef.Occurrences.Count,
                                total_constraints = asmDef.Constraints.Count
                            }
                        };

                        Console.WriteLine("Extracting components...");
                        ExtractOccurrences(asmDef.Occurrences, export.components, "");

                        Console.WriteLine("Extracting constraints...");
                        ExtractConstraints(asmDef.Constraints, export.constraints);

                        string fileName = System.IO.Path.GetFileNameWithoutExtension(assemblyPath);
                        string outputPath = System.IO.Path.Combine(outputRoot, fileName + "_full_export.json");

                        System.IO.File.WriteAllText(outputPath,
                            JsonConvert.SerializeObject(export, Formatting.Indented));

                        Console.WriteLine("Export complete -> " + outputPath);

                        asmDoc.Close(true);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Failed: " + assemblyPath);
                        Console.WriteLine(ex.Message);
                    }
                }
            }

            Console.WriteLine("All assemblies processed.");
        }

        // ════════════════════════════════════════════════════════════════════
        // OCCURRENCES  (recursive)
        // ════════════════════════════════════════════════════════════════════

        static void ExtractOccurrences(
            ComponentOccurrences occurrences,
            List<ComponentData> list,
            string parentPath)
        {
            foreach (ComponentOccurrence occ in occurrences)
            {
                Matrix m = occ.Transformation;

                double[][] rot = new double[3][];
                for (int i = 0; i < 3; i++)
                {
                    rot[i] = new double[3];
                    for (int j = 0; j < 3; j++)
                        rot[i][j] = m.get_Cell(i + 1, j + 1);
                }

                ComponentData comp = new ComponentData
                {
                    occurrence_name = occ.Name,
                    occurrence_path = parentPath == "" ? occ.Name : parentPath + "/" + occ.Name,
                    file_name = occ.Definition.Document.FullFileName,
                    component_type = occ.DefinitionDocumentType.ToString(),
                    grounded = occ.Grounded,
                    suppressed = occ.Suppressed,
                    visible = occ.Visible,
                    transform = new TransformData
                    {
                        rotation_matrix = rot,
                        translation_cm = new double[]
                        {
                            m.get_Cell(1, 4),
                            m.get_Cell(2, 4),
                            m.get_Cell(3, 4)
                        }
                    }
                };

                list.Add(comp);

                if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    try
                    {
                        AssemblyComponentDefinition subDef =
                            (AssemblyComponentDefinition)occ.Definition;
                        if (subDef.Occurrences.Count > 0)
                            ExtractOccurrences(subDef.Occurrences,
                                               comp.sub_components,
                                               comp.occurrence_path);
                    }
                    catch { }
                }
            }
        }

        // ════════════════════════════════════════════════════════════════════
        // CONSTRAINTS
        // ════════════════════════════════════════════════════════════════════

        static void ExtractConstraints(
            AssemblyConstraints constraints,
            List<ConstraintData> list)
        {
            foreach (AssemblyConstraint c in constraints)
            {
                ConstraintData con = new ConstraintData
                {
                    constraint_name = c.Name,
                    constraint_type = c.Type.ToString(),
                    suppressed = c.Suppressed,
                    occurrence_one = c.OccurrenceOne != null ? c.OccurrenceOne.Name : null,
                    occurrence_two = c.OccurrenceTwo != null ? c.OccurrenceTwo.Name : null
                };

                if (c is MateConstraint)
                {
                    MateConstraint mate = (MateConstraint)c;
                    if (mate.Offset != null) con.offset_cm = mate.Offset.Value;
                    con.mate_type = mate.Type.ToString();
                }
                else if (c is FlushConstraint)
                {
                    FlushConstraint flush = (FlushConstraint)c;
                    if (flush.Offset != null) con.offset_cm = flush.Offset.Value;
                }
                else if (c is AngleConstraint)
                {
                    AngleConstraint angle = (AngleConstraint)c;
                    if (angle.Angle != null) con.angle_rad = angle.Angle.Value;
                    con.angle_type = angle.Type.ToString();
                }
                else if (c is TangentConstraint)
                {
                    TangentConstraint tangent = (TangentConstraint)c;
                    con.inside_tangency = tangent.InsideTangency;
                }
                else if (c is InsertConstraint)
                {
                    con.constraint_extra_info = "InsertConstraint";
                }
                else if (c is SymmetryConstraint)
                {
                    con.constraint_extra_info = "SymmetryConstraint";
                }
                else
                {
                    con.constraint_extra_info = c.Type.ToString();
                }

                con.entity_one = ExtractEntityData(c.EntityOne);
                con.entity_two = ExtractEntityData(c.EntityTwo);

                list.Add(con);
            }
        }

        // ════════════════════════════════════════════════════════════════════
        // ENTITY — reference key + geometric signature
        // ════════════════════════════════════════════════════════════════════

        static EntityData ExtractEntityData(object entityObj)
        {
            if (entityObj == null) return null;

            EntityData data = new EntityData();
            object native = null;

            if (entityObj is FaceProxy)
            {
                FaceProxy fp = (FaceProxy)entityObj;
                data.entity_type = "FaceProxy";
                native = fp.NativeObject;
                data.proxy_context_occurrence =
                    fp.ContainingOccurrence != null ? fp.ContainingOccurrence.Name : null;
                try { Face f = (Face)native; data.surface_type = f.SurfaceType.ToString(); data.area_cm2 = f.Evaluator.Area; } catch { }
            }
            else if (entityObj is EdgeProxy)
            {
                EdgeProxy ep = (EdgeProxy)entityObj;
                data.entity_type = "EdgeProxy";
                native = ep.NativeObject;
                data.proxy_context_occurrence =
                    ep.ContainingOccurrence != null ? ep.ContainingOccurrence.Name : null;
                try { Edge e = (Edge)native; data.curve_type = e.CurveType.ToString(); data.length_cm = GetEdgeLength(e); } catch { }
            }
            else if (entityObj is WorkPlaneProxy)
            {
                WorkPlaneProxy wpp = (WorkPlaneProxy)entityObj;
                data.entity_type = "WorkPlaneProxy";
                native = wpp.NativeObject;
                data.proxy_context_occurrence =
                    wpp.ContainingOccurrence != null ? wpp.ContainingOccurrence.Name : null;
                try { data.work_feature_name = ((WorkPlane)native).Name; } catch { }
            }
            else if (entityObj is WorkAxisProxy)
            {
                WorkAxisProxy wap = (WorkAxisProxy)entityObj;
                data.entity_type = "WorkAxisProxy";
                native = wap.NativeObject;
                data.proxy_context_occurrence =
                    wap.ContainingOccurrence != null ? wap.ContainingOccurrence.Name : null;
                try { data.work_feature_name = ((WorkAxis)native).Name; } catch { }
            }
            else if (entityObj is WorkPointProxy)
            {
                WorkPointProxy wptp = (WorkPointProxy)entityObj;
                data.entity_type = "WorkPointProxy";
                native = wptp.NativeObject;
                data.proxy_context_occurrence =
                    wptp.ContainingOccurrence != null ? wptp.ContainingOccurrence.Name : null;
                try { data.work_feature_name = ((WorkPoint)native).Name; } catch { }
            }
            else if (entityObj is VertexProxy)
            {
                VertexProxy vp = (VertexProxy)entityObj;
                data.entity_type = "VertexProxy";
                native = vp.NativeObject;
                data.proxy_context_occurrence =
                    vp.ContainingOccurrence != null ? vp.ContainingOccurrence.Name : null;
            }
            else if (entityObj is Face)
            {
                Face f = (Face)entityObj;
                data.entity_type = "Face";
                native = f;
                try { data.surface_type = f.SurfaceType.ToString(); data.area_cm2 = f.Evaluator.Area; } catch { }
            }
            else if (entityObj is Edge)
            {
                Edge e = (Edge)entityObj;
                data.entity_type = "Edge";
                native = e;
                try { data.curve_type = e.CurveType.ToString(); data.length_cm = GetEdgeLength(e); } catch { }
            }
            else
            {
                data.entity_type = entityObj.GetType().Name;
                return data;
            }

            if (native == null) return data;

            // ── Reference key + round-trip validation ────────────────────────
            try
            {
                Document doc = GetDocumentFromNative(native);
                if (doc == null) return data;

                data.owner_document = doc.FullFileName;

                ReferenceKeyManager mgr = doc.ReferenceKeyManager;
                int keyContext = mgr.CreateKeyContext();

                byte[] referenceKey = new byte[1];
                CallGetReferenceKey(native, ref referenceKey, keyContext);

                // SaveContextToArray  →  arg 2 is 'ref'
                byte[] contextKeyArray = new byte[1];
                mgr.SaveContextToArray(keyContext, ref contextKeyArray);

                string refKeyString = mgr.KeyToString(referenceKey);
                string ctxKeyString = mgr.KeyToString(contextKeyArray);

                data.reference_key_string = refKeyString;
                data.context_key_string = ctxKeyString;

                // StringToKey  →  arg 2 is 'ref'
                byte[] referenceKeyAfter = new byte[1];
                mgr.StringToKey(refKeyString, ref referenceKeyAfter);

                byte[] contextKeyAfter = new byte[1];
                mgr.StringToKey(ctxKeyString, ref contextKeyAfter);

                // LoadContextFromArray  →  arg 1 is 'ref'
                int restoredContext = mgr.LoadContextFromArray(ref contextKeyAfter);

                object matchType;
                object boundEntity = mgr.BindKeyToObject(referenceKeyAfter, restoredContext, out matchType);

                data.bind_match_type = ((SolutionNatureEnum)matchType).ToString();
                data.bind_succeeded = (boundEntity != null);
            }
            catch (Exception ex)
            {
                data.key_extraction_error = ex.Message;
            }

            // ── Face geometric signature ─────────────────────────────────────
            //
            // VERIFIED from Autodesk community working code:
            //
            //   SurfaceEvaluator.ParamRangeRect  →  returns Box2d  (UV bounding rect)
            //
            //   SurfaceEvaluator.GetParamAtPoint(
            //       ref double[] point,            // [x,y,z]  input  3 elements
            //       ref double[] guessParam,       // [u,v]    hint   2 elements (zeros ok)
            //       ref double[] maxDeviation,     // [u,v]    output 2 elements
            //       ref double[] param,            // [u,v]    output 2 elements
            //       ref SolutionNatureEnum[] sol)  //          output 1 element
            //
            //   SurfaceEvaluator.GetNormal(
            //       ref double[] params,           // [u,v]    input  2 elements
            //       ref double[] normals)          // [x,y,z]  output 3 elements
            //                                      // * 2 args only — no points output *
            //
            if (native is Face)
            {
                try
                {
                    Face face = (Face)native;
                    SurfaceEvaluator eval = face.Evaluator;

                    // Use ParamRangeRect (Box2d) for UV extents — not RangeBox
                    Box2d uvRange = eval.ParamRangeRect;

                    // Bounding box in model space for size fingerprint
                    Box box = eval.RangeBox;
                    if (box != null)
                    {
                        data.face_bbox_min = new double[] { box.MinPoint.X, box.MinPoint.Y, box.MinPoint.Z };
                        data.face_bbox_max = new double[] { box.MaxPoint.X, box.MaxPoint.Y, box.MaxPoint.Z };
                    }

                    // Build UV midpoint from ParamRangeRect
                    double[] pars = new double[2];
                    pars[0] = (uvRange.MinPoint.X + uvRange.MaxPoint.X) / 2.0;
                    pars[1] = (uvRange.MinPoint.Y + uvRange.MaxPoint.Y) / 2.0;

                    // Sample normal at UV midpoint — GetNormal takes exactly 2 args
                    double[] normals = new double[3];
                    eval.GetNormal(ref pars, ref normals);

                    data.face_normal_at_center = normals;

                    // Also get the 3D point on the surface at this UV using GetPointAtParam
                    // (reuses the pars array from above as input)
                    double[] surfPt = new double[3];
                    eval.GetPointAtParam(ref pars, ref surfPt);
                    data.face_point_at_center = surfPt;
                }
                catch { }
            }

            // ── Edge geometric signature ─────────────────────────────────────
            //
            // VERIFIED Inventor COM signatures:
            //
            //   CurveEvaluator.GetParamExtents(out double start, out double end)
            //
            //   CurveEvaluator.GetPointAtParam(
            //       ref double[] params,    // input  1 element
            //       ref double[] points)    // output 3 elements  ← 'ref' NOT 'out'
            //
            //   CurveEvaluator.GetTangent(
            //       ref double[] params,    // input  1 element
            //       ref double[] tangents)  // output 3 elements  ← 'ref' NOT 'out'
            //
            if (native is Edge)
            {
                try
                {
                    Edge edge = (Edge)native;
                    CurveEvaluator eval = edge.Evaluator;   // type is CurveEvaluator

                    // GetParamExtents: both args are 'out'
                    double startParam, endParam;
                    eval.GetParamExtents(out startParam, out endParam);

                    double mid = (startParam + endParam) / 2.0;
                    double[] pArr = new double[] { mid };

                    // GetPointAtParam: arg 2 is 'ref'
                    double[] midPoint = new double[3];
                    eval.GetPointAtParam(ref pArr, ref midPoint);

                    // GetTangent: arg 2 is 'ref'
                    double[] tangent = new double[3];
                    eval.GetTangent(ref pArr, ref tangent);

                    data.edge_midpoint = midPoint;
                    data.edge_tangent_at_mid = tangent;
                }
                catch { }
            }

            return data;
        }

        // ════════════════════════════════════════════════════════════════════
        // EDGE LENGTH
        // ════════════════════════════════════════════════════════════════════

        static double GetEdgeLength(Edge edge)
        {
            try
            {
                CurveEvaluator eval = edge.Evaluator;   // CurveEvaluator (not CurveEvaluator3D)

                if (edge.CurveType == CurveTypeEnum.kLineCurve)
                {
                    Point p1 = edge.StartVertex.Point;
                    Point p2 = edge.StopVertex.Point;
                    double dx = p2.X - p1.X;
                    double dy = p2.Y - p1.Y;
                    double dz = p2.Z - p1.Z;
                    return Math.Sqrt(dx * dx + dy * dy + dz * dz);
                }
                else if (edge.CurveType == CurveTypeEnum.kCircleCurve)
                {
                    Circle circ = (Circle)edge.Geometry;
                    double radius = circ.Radius;
                    double start, end;
                    eval.GetParamExtents(out start, out end);   // 'out'
                    return radius * Math.Abs(end - start);
                }
                else
                {
                    double start, end;
                    eval.GetParamExtents(out start, out end);   // 'out'
                    return Math.Abs(end - start);
                }
            }
            catch
            {
                return 0.0;
            }
        }

        // ════════════════════════════════════════════════════════════════════
        // HELPERS
        // ════════════════════════════════════════════════════════════════════

        static Document GetDocumentFromNative(object native)
        {
            try
            {
                if (native is Face) return (Document)((Face)native).Parent.ComponentDefinition.Document;
                if (native is Edge) return (Document)((Edge)native).Parent.ComponentDefinition.Document;
                if (native is WorkPlane) return (Document)((WorkPlane)native).Parent.Document;
                if (native is WorkAxis) return (Document)((WorkAxis)native).Parent.Document;
                if (native is WorkPoint) return (Document)((WorkPoint)native).Parent.Document;
                if (native is Vertex) return (Document)((Vertex)native).Parent.ComponentDefinition.Document;
            }
            catch { }
            return null;
        }

        static void CallGetReferenceKey(object native, ref byte[] key, int context)
        {
            if (native is Face) { ((Face)native).GetReferenceKey(ref key, context); return; }
            if (native is Edge) { ((Edge)native).GetReferenceKey(ref key, context); return; }
            if (native is WorkPlane) { ((WorkPlane)native).GetReferenceKey(ref key, context); return; }
            if (native is WorkAxis) { ((WorkAxis)native).GetReferenceKey(ref key, context); return; }
            if (native is WorkPoint) { ((WorkPoint)native).GetReferenceKey(ref key, context); return; }
            if (native is Vertex) { ((Vertex)native).GetReferenceKey(ref key, context); return; }
            throw new NotSupportedException("GetReferenceKey not implemented for: " + native.GetType().Name);
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    // DATA STRUCTURES
    // ══════════════════════════════════════════════════════════════════════

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
        // ── Identity ──────────────────────────────────────────────────────
        public string entity_type { get; set; }
        public string proxy_context_occurrence { get; set; }
        public string owner_document { get; set; }

        // ── Reference keys ───────────────────────────────────────────────
        public string reference_key_string { get; set; }
        public string context_key_string { get; set; }
        public bool? bind_succeeded { get; set; }
        public string bind_match_type { get; set; }
        public string key_extraction_error { get; set; }

        // ── Face signature ────────────────────────────────────────────────
        public string surface_type { get; set; }
        public double? area_cm2 { get; set; }
        public double[] face_normal_at_center { get; set; }   // unit normal at UV midpoint
        public double[] face_point_at_center { get; set; }   // 3D point at UV midpoint
        public double[] face_bbox_min { get; set; }   // model-space bbox min
        public double[] face_bbox_max { get; set; }   // model-space bbox max

        // ── Edge signature ────────────────────────────────────────────────
        public string curve_type { get; set; }
        public double? length_cm { get; set; }
        public double[] edge_midpoint { get; set; }   // 3D point at mid-param
        public double[] edge_tangent_at_mid { get; set; }   // tangent at mid-param

        // ── Work feature ──────────────────────────────────────────────────
        public string work_feature_name { get; set; }
    }
}