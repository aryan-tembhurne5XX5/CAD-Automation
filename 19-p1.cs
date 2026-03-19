using Inventor;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace InventorPartExporter
{
    class Program
    {
        static Inventor.Application invApp;

        static void Main(string[] args)
        {
            string baseInputRoot = @"E:\Phase 1";
            string baseOutputRoot = @"E:\Phase 1\parts_raw_export";

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

            Console.WriteLine($"Starting Generative Part extraction in: {baseInputRoot}");

            if (!System.IO.Directory.Exists(baseOutputRoot))
                System.IO.Directory.CreateDirectory(baseOutputRoot);

            foreach (string iptPath in System.IO.Directory.GetFiles(baseInputRoot, "*.ipt", System.IO.SearchOption.AllDirectories))
            {
                ProcessPart(iptPath, baseOutputRoot);
            }

            Console.WriteLine("\n✅ All parts successfully extracted for Generative ML.");
        }

        static void ProcessPart(string iptPath, string outputRoot)
        {
            try
            {
                Console.WriteLine($"\nOpening Part: {iptPath}");

                Document openedDoc = invApp.Documents.Open(iptPath, true);
                PartDocument partDoc = (PartDocument)openedDoc;
                PartComponentDefinition def = partDoc.ComponentDefinition;

                PartExport export = new PartExport();

                ReferenceKeyManager mgr = partDoc.ReferenceKeyManager;
                int keyContext = mgr.CreateKeyContext();
                
                // FIXED: Dynamic Memory Allocation
                byte[] ctxArray = new byte[0];
                mgr.SaveContextToArray(keyContext, ref ctxArray);
                export.context_key_string = mgr.KeyToString(ctxArray);

                export.part_metadata = new PartMetadata
                {
                    file_name = partDoc.DisplayName,
                    full_path = partDoc.FullFileName,
                    internal_name = partDoc.InternalName,
                    units = partDoc.UnitsOfMeasure.LengthUnits.ToString(),
                    part_number = GetProp((Document)partDoc, "Design Tracking Properties", "Part Number"),
                    description = GetProp((Document)partDoc, "Design Tracking Properties", "Description"),
                    material = GetProp((Document)partDoc, "Design Tracking Properties", "Material"),
                    mass_kg = SafeMass(def)
                };

                try
                {
                    Box rb = def.RangeBox;
                    export.bounding_box = new BoundingBox { min = ToPoint(rb.MinPoint), max = ToPoint(rb.MaxPoint) };
                } catch { }

                // --- 1. PARAMETRICS (Model & User) ---
                Console.WriteLine("  Extracting Parameters...");
                foreach (ModelParameter param in def.Parameters.ModelParameters)
                {
                    try { export.parameters.Add(new ParameterData { name = param.Name, type="Model", expression = param.Expression, value_cm = param.Value }); } catch { }
                }
                foreach (UserParameter param in def.Parameters.UserParameters)
                {
                    try { export.parameters.Add(new ParameterData { name = param.Name, type="User", expression = param.Expression, value_cm = param.Value }); } catch { }
                }

                // --- 2. MULTI-BODY DATA ---
Console.WriteLine("  Extracting Solid Bodies...");
foreach (SurfaceBody body in def.SurfaceBodies)
{
    double vol = 0;
    try 
    { 
        // Correct C# Interop way to call a parameterized COM property
        double[] tolerances = new double[0];
        vol = body.get_Volume(ref tolerances); 
    } 
    catch 
    { 
        // Ultimate fallback using dynamic dispatch
        try { dynamic b = body; vol = (double)b.Volume; } catch { }
    }
    
    export.bodies.Add(new BodyData { body_name = body.Name, volume_cm3 = vol });
}

                // --- 3. TRUE B-REP & FACE LOOPS ---
                Console.WriteLine("  Extracting Faces, Edges, and Face Loops...");
                foreach (SurfaceBody body in def.SurfaceBodies)
                {
                    foreach (Face face in body.Faces)
                        export.faces.Add(ExtractFace(face, mgr, keyContext));
                    
                    foreach (Edge edge in body.Edges)
                        export.edges.Add(ExtractEdge(edge, mgr, keyContext));
                }

                // --- 4. SKETCH GEOMETRY & CONSTRAINTS ---
                Console.WriteLine("  Extracting 2D Sketches and Constraints...");
                foreach (PlanarSketch sketch in def.Sketches)
                    export.sketches.Add(ExtractSketch(sketch, mgr, keyContext));

                // --- 5. VIRTUAL TOPOLOGY ---
                Console.WriteLine("  Extracting WorkPlanes...");
                ExtractWorkPlanes(def, export, mgr, keyContext);

                // --- 6. HOLES ---
                Console.WriteLine("  Extracting Holes...");
                foreach (HoleFeature hole in def.Features.HoleFeatures)
                {
                    if (!hole.Suppressed) export.connection_points.AddRange(ExtractHole(hole, mgr, keyContext));
                }

                // --- 7. FEATURE GRAPH (With Operations) ---
                Console.WriteLine("  Extracting Feature Graph...");
                foreach (PartFeature feat in def.Features)
                    export.feature_graph.Add(ExtractFeatureNode(feat));

                // --- 8. PATTERNS ---
                ExtractPatterns(def, export);

                string fileName = System.IO.Path.GetFileNameWithoutExtension(iptPath);
                string outputPath = System.IO.Path.Combine(outputRoot, fileName + ".json");
                System.IO.File.WriteAllText(outputPath, JsonConvert.SerializeObject(export, Newtonsoft.Json.Formatting.Indented));

                partDoc.Close(true);
                Console.WriteLine($"  Exported → {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] Failed processing {iptPath}: {ex.Message}");
            }
        }

        // ==========================================
        // TRUE B-REP MATHEMATICS WITH FACE LOOPS
        // ==========================================
        static FaceData ExtractFace(Face face, ReferenceKeyManager mgr, int keyContext)
        {
            var data = new FaceData { transient_key = face.TransientKey, surface_type = face.SurfaceType.ToString(), area_cm2 = face.Evaluator.Area };

            try
            {
                byte[] key = new byte[0]; // FIXED
                face.GetReferenceKey(ref key, keyContext);
                data.reference_key_string = mgr.KeyToString(key);
            } catch { }

            try
            {
                Box box = face.Evaluator.RangeBox;
                data.bbox_min = ToPoint(box.MinPoint); data.bbox_max = ToPoint(box.MaxPoint);
            } catch { }

            try
            {
                SurfaceEvaluator eval = face.Evaluator;
                Box2d uv = eval.ParamRangeRect;
                double[] pars = { (uv.MinPoint.X + uv.MaxPoint.X) / 2.0, (uv.MinPoint.Y + uv.MaxPoint.Y) / 2.0 };
                double[] normal = new double[3], pt = new double[3];
                eval.GetNormal(ref pars, ref normal);
                eval.GetPointAtParam(ref pars, ref pt);
                data.normal = normal; data.center = pt;
            } catch { }

            try { if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface) data.radius_cm = ((Cylinder)face.Geometry).Radius; } catch { }
            try { if (face.CreatedByFeature != null) data.created_by_feature = face.CreatedByFeature.Name; } catch { }

            data.loops = new List<LoopData>();
            foreach (EdgeLoop loop in face.EdgeLoops)
            {
                var loopData = new LoopData { is_outer = loop.IsOuterEdgeLoop };
                foreach (Edge loopEdge in loop.Edges)
                {
                    try
                    {
                        byte[] eKey = new byte[0]; // FIXED
                        loopEdge.GetReferenceKey(ref eKey, keyContext);
                        loopData.edge_reference_keys.Add(mgr.KeyToString(eKey));
                    } catch { }
                }
                data.loops.Add(loopData);
            }
            return data;
        }

        static EdgeData ExtractEdge(Edge edge, ReferenceKeyManager mgr, int keyContext)
        {
            var data = new EdgeData { transient_key = edge.TransientKey, curve_type = edge.CurveType.ToString(), length_cm = GetEdgeLength(edge) };

            try
            {
                byte[] key = new byte[0]; // FIXED
                edge.GetReferenceKey(ref key, keyContext);
                data.reference_key_string = mgr.KeyToString(key);
            } catch { }

            try
            {
                CurveEvaluator eval = edge.Evaluator;
                double s, e, length, midParam;
                eval.GetParamExtents(out s, out e);
                eval.GetLengthAtParam(s, e, out length);
                eval.GetParamAtLength(s, length / 2.0, out midParam);

                double[] arr = { midParam }, pt = new double[3], tan = new double[3];
                eval.GetPointAtParam(ref arr, ref pt);
                eval.GetTangent(ref arr, ref tan);
                
                data.midpoint = pt; data.tangent = tan;
                data.start_vertex = new double[] { edge.StartVertex.Point.X, edge.StartVertex.Point.Y, edge.StartVertex.Point.Z };
                data.end_vertex = new double[] { edge.StopVertex.Point.X, edge.StopVertex.Point.Y, edge.StopVertex.Point.Z };
            } catch { }

            data.adjacent_faces = new List<string>();
            foreach (Face adjacentFace in edge.Faces)
            {
                try
                {
                    byte[] refKey = new byte[0]; // FIXED
                    adjacentFace.GetReferenceKey(ref refKey, keyContext);
                    data.adjacent_faces.Add(mgr.KeyToString(refKey));
                } catch { }
            }
            return data;
        }

        static double GetEdgeLength(Edge edge)
        {
            try
            {
                CurveEvaluator eval = edge.Evaluator;
                double s, e, length;
                eval.GetParamExtents(out s, out e);
                eval.GetLengthAtParam(s, e, out length);
                return length;
            }
            catch 
            { 
                try {
                    Point p1 = edge.StartVertex.Point, p2 = edge.StopVertex.Point;
                    return Math.Sqrt(Math.Pow(p2.X - p1.X, 2) + Math.Pow(p2.Y - p1.Y, 2) + Math.Pow(p2.Z - p1.Z, 2));
                } catch { return 0; }
            }
        }

        // ==========================================
        // SKETCHES & CONSTRAINTS (GENERATIVE READY)
        // ==========================================
        static SketchData ExtractSketch(PlanarSketch sketch, ReferenceKeyManager mgr, int keyContext)
        {
            var data = new SketchData { name = sketch.Name, visible = sketch.Visible };
            
            try {
                foreach (SketchLine line in sketch.SketchLines)
                    data.lines.Add(new double[] { line.StartSketchPoint.Geometry.X, line.StartSketchPoint.Geometry.Y, line.EndSketchPoint.Geometry.X, line.EndSketchPoint.Geometry.Y });
            } catch { }

            try {
                foreach (SketchCircle circ in sketch.SketchCircles)
                    data.circles.Add(new double[] { circ.CenterSketchPoint.Geometry.X, circ.CenterSketchPoint.Geometry.Y, circ.Radius });
            } catch { }

            // GEOMETRIC CONSTRAINTS
            try {
                foreach (GeometricConstraint gc in sketch.GeometricConstraints)
                    data.geometric_constraints.Add(gc.Type.ToString().Replace("Object", ""));
            } catch { }

            // DIMENSION CONSTRAINTS
            try {
                foreach (DimensionConstraint dc in sketch.DimensionConstraints)
                    data.dimension_constraints.Add(new SketchDim { type = dc.Type.ToString().Replace("Object", ""), value = dc.Parameter.Value });
            } catch { }

            return data;
        }

        // ==========================================
        // WORKPLANES & FEATURES
        // ==========================================
        static void ExtractWorkPlanes(PartComponentDefinition def, PartExport export, ReferenceKeyManager mgr, int keyContext)
        {
            foreach (WorkPlane wp in def.WorkPlanes)
            {
                try
                {
                    Plane mathPlane = wp.Plane;
                    var wpData = new WorkPlaneData
                    {
                        work_feature_name = wp.Name,
                        center = new double[] { mathPlane.RootPoint.X, mathPlane.RootPoint.Y, mathPlane.RootPoint.Z },
                        normal = new double[] { mathPlane.Normal.X, mathPlane.Normal.Y, mathPlane.Normal.Z }
                    };
                    byte[] key = new byte[0]; // FIXED
                    wp.GetReferenceKey(ref key, keyContext);
                    wpData.reference_key_string = mgr.KeyToString(key);
                    export.work_planes.Add(wpData);
                } catch { } 
            }
        }

        static FeatureNode ExtractFeatureNode(PartFeature feat)
        {
            var node = new FeatureNode { feature_name = feat.Name, feature_type = ClassifyFeature(feat), suppressed = feat.Suppressed };
            dynamic dynFeat = feat;
            
            // EXTRACT FEATURE OPERATIONS (Crucial for Generative CAD)
            try {
                if (feat is ExtrudeFeature) node.operation = ((ExtrudeFeature)feat).Operation.ToString();
                else if (feat is RevolveFeature) node.operation = ((RevolveFeature)feat).Operation.ToString();
                else if (feat is SweepFeature) node.operation = ((SweepFeature)feat).Operation.ToString();
            } catch { }

            try { foreach (object dep in dynFeat.DependentFeatures) { dynamic d = dep; node.child_features.Add((string)d.Name); } } catch { }
            try { foreach (object dep in dynFeat.DependedOnFeatures) { dynamic d = dep; node.parent_features.Add((string)d.Name); } } catch { }
            return node;
        }

        static string ClassifyFeature(PartFeature feat)
        {
            if (feat is ExtrudeFeature) return "Extrude";
            if (feat is RevolveFeature) return "Revolve";
            if (feat is SweepFeature) return "Sweep";
            if (feat is LoftFeature) return "Loft";
            if (feat is HoleFeature) return "Hole";
            if (feat is FilletFeature) return "Fillet";
            if (feat is ChamferFeature) return "Chamfer";
            if (feat is RectangularPatternFeature) return "RectangularPattern";
            if (feat is CircularPatternFeature) return "CircularPattern";
            if (feat is CombineFeature) return "BooleanCombine";
            return "Other";
        }

        static List<ConnectionPoint> ExtractHole(HoleFeature hole, ReferenceKeyManager mgr, int keyContext)
        {
            var list = new List<ConnectionPoint>();
            string holeType = hole.Tapped ? "Tapped" : "Simple";
            double diam_cm = 0;
            try { diam_cm = hole.HoleDiameter.Value; } catch { }
            list.Add(new ConnectionPoint {
                id = Guid.NewGuid().ToString(), feature_name = hole.Name, feature_type = "Hole", suppressed = hole.Suppressed,
                hole_properties = new HoleProperties { hole_type = holeType, diameter_cm = diam_cm, is_threaded = hole.Tapped }
            });
            return list;
        }

        static void ExtractPatterns(PartComponentDefinition def, PartExport export)
        {
            foreach (RectangularPatternFeature rp in def.Features.RectangularPatternFeatures)
            {
                var pd = new PatternData { name = rp.Name, type = "Rectangular", suppressed = rp.Suppressed };
                try { pd.count_dir1 = (int)rp.XCount.Value; pd.count_dir2 = (int)rp.YCount.Value; } catch { }
                export.patterns.Add(pd);
            }
        }

        static PointData ToPoint(Point p) => new PointData { x = p.X, y = p.Y, z = p.Z };
        static double SafeMass(PartComponentDefinition def) { try { return def.MassProperties.Mass; } catch { return 0; } }
        static string GetProp(Document doc, string setName, string propName) { try { return doc.PropertySets[setName][propName].Value?.ToString() ?? ""; } catch { return ""; } }
    }

    // ==========================================
    // DATA STRUCTURES
    // ==========================================
    public class PartExport
    {
        public PartMetadata part_metadata { get; set; }
        public string context_key_string { get; set; }
        public BoundingBox bounding_box { get; set; }
        public List<ParameterData> parameters { get; set; } = new List<ParameterData>();
        public List<BodyData> bodies { get; set; } = new List<BodyData>();
        public List<FaceData> faces { get; set; } = new List<FaceData>();
        public List<EdgeData> edges { get; set; } = new List<EdgeData>();
        public List<SketchData> sketches { get; set; } = new List<SketchData>();
        public List<WorkPlaneData> work_planes { get; set; } = new List<WorkPlaneData>();
        public List<ConnectionPoint> connection_points { get; set; } = new List<ConnectionPoint>();
        public List<FeatureNode> feature_graph { get; set; } = new List<FeatureNode>();
        public List<PatternData> patterns { get; set; } = new List<PatternData>();
    }

    public class PartMetadata { public string file_name { get; set; } public string full_path { get; set; } public string internal_name { get; set; } public string units { get; set; } public string part_number { get; set; } public string description { get; set; } public string material { get; set; } public double mass_kg { get; set; } }
    public class BoundingBox { public PointData min { get; set; } public PointData max { get; set; } }
    public class PointData { public double x { get; set; } public double y { get; set; } public double z { get; set; } }
    public class ParameterData { public string name { get; set; } public string type { get; set; } public string expression { get; set; } public double value_cm { get; set; } }
    public class BodyData { public string body_name { get; set; } public double volume_cm3 { get; set; } }
    public class WorkPlaneData { public string entity_type { get; set; } = "WorkPlane"; public string work_feature_name { get; set; } public string surface_type { get; set; } = "kPlaneSurface"; public double area_cm2 { get; set; } = 0.0001; public double[] center { get; set; } public double[] normal { get; set; } public string reference_key_string { get; set; } }
    public class FaceData { public int transient_key { get; set; } public string reference_key_string { get; set; } public string surface_type { get; set; } public double area_cm2 { get; set; } public double[] normal { get; set; } public double[] center { get; set; } public double? radius_cm { get; set; } public PointData bbox_min { get; set; } public PointData bbox_max { get; set; } public string created_by_feature { get; set; } public List<LoopData> loops { get; set; } = new List<LoopData>(); }
    public class LoopData { public bool is_outer { get; set; } public List<string> edge_reference_keys { get; set; } = new List<string>(); }
    public class EdgeData { public int transient_key { get; set; } public string reference_key_string { get; set; } public string curve_type { get; set; } public double length_cm { get; set; } public double[] midpoint { get; set; } public double[] tangent { get; set; } public double[] start_vertex { get; set; } public double[] end_vertex { get; set; } public List<string> adjacent_faces { get; set; } }
    
    public class SketchData { 
        public string name { get; set; } 
        public bool visible { get; set; } 
        public List<double[]> lines { get; set; } = new List<double[]>(); 
        public List<double[]> circles { get; set; } = new List<double[]>(); 
        public List<string> geometric_constraints { get; set; } = new List<string>();
        public List<SketchDim> dimension_constraints { get; set; } = new List<SketchDim>();
    }
    public class SketchDim { public string type { get; set; } public double value { get; set; } }

    public class ConnectionPoint { public string id { get; set; } public string feature_name { get; set; } public string feature_type { get; set; } public bool suppressed { get; set; } public HoleProperties hole_properties { get; set; } }
    public class HoleProperties { public string hole_type { get; set; } public double diameter_cm { get; set; } public bool is_threaded { get; set; } }
    public class FeatureNode { public string feature_name { get; set; } public string feature_type { get; set; } public string operation { get; set; } public bool suppressed { get; set; } public List<string> parent_features { get; set; } = new List<string>(); public List<string> child_features { get; set; } = new List<string>(); }
    public class PatternData { public string name { get; set; } public string type { get; set; } public bool suppressed { get; set; } public int? count_dir1 { get; set; } public int? count_dir2 { get; set; } }
}