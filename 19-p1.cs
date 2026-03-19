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
            // UPDATE THESE PATHS TO YOUR FOLDERS
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

            Console.WriteLine($"Starting Full Recursive Part extraction in: {baseInputRoot}");

            if (!System.IO.Directory.Exists(baseOutputRoot))
                System.IO.Directory.CreateDirectory(baseOutputRoot);

            // Recursively find ALL .ipt files
            foreach (string iptPath in System.IO.Directory.GetFiles(baseInputRoot, "*.ipt", System.IO.SearchOption.AllDirectories))
            {
                ProcessPart(iptPath, baseOutputRoot);
            }

            Console.WriteLine("\n✅ All parts successfully extracted.");
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
                byte[] ctxArray = new byte[1];
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
                    export.bounding_box = new BoundingBox
                    {
                        min = ToPoint(rb.MinPoint),
                        max = ToPoint(rb.MaxPoint)
                    };
                }
                catch { }

                Console.WriteLine("  Extracting faces and edges...");
                foreach (SurfaceBody body in def.SurfaceBodies)
                {
                    foreach (Face face in body.Faces)
                        export.faces.Add(ExtractFace(face, mgr, keyContext));
                    
                    foreach (Edge edge in body.Edges)
                        export.edges.Add(ExtractEdge(edge, mgr, keyContext));
                }

                Console.WriteLine("  Extracting WorkPlanes (Virtual Topology)...");
                ExtractWorkPlanes(def, export, mgr, keyContext);

                // --- SAVE TO JSON ---
                string fileName = System.IO.Path.GetFileNameWithoutExtension(iptPath);
                string outputPath = System.IO.Path.Combine(outputRoot, fileName + ".json");

                string json = JsonConvert.SerializeObject(export, Newtonsoft.Json.Formatting.Indented);
                System.IO.File.WriteAllText(outputPath, json);

                partDoc.Close(true);
                Console.WriteLine($"  Exported → {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] Failed processing {iptPath}: {ex.Message}");
            }
        }

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

                    byte[] key = new byte[1];
                    wp.GetReferenceKey(ref key, keyContext);
                    wpData.reference_key_string = mgr.KeyToString(key);
                    
                    export.work_planes.Add(wpData);
                }
                catch { } 
            }
        }

        static FaceData ExtractFace(Face face, ReferenceKeyManager mgr, int keyContext)
        {
            var data = new FaceData
            {
                transient_key = face.TransientKey,
                surface_type = face.SurfaceType.ToString(),
                area_cm2 = face.Evaluator.Area
            };

            try
            {
                byte[] key = new byte[1];
                face.GetReferenceKey(ref key, keyContext);
                data.reference_key_string = mgr.KeyToString(key);
            }
            catch { }

            try
            {
                Box box = face.Evaluator.RangeBox;
                data.bbox_min = ToPoint(box.MinPoint);
                data.bbox_max = ToPoint(box.MaxPoint);
            }
            catch { }

            try
            {
                SurfaceEvaluator eval = face.Evaluator;
                Box2d uv = eval.ParamRangeRect;
                double[] pars = { (uv.MinPoint.X + uv.MaxPoint.X) / 2.0, (uv.MinPoint.Y + uv.MaxPoint.Y) / 2.0 };
                double[] normal = new double[3];
                double[] pt = new double[3];
                eval.GetNormal(ref pars, ref normal);
                eval.GetPointAtParam(ref pars, ref pt);
                data.normal = normal;
                data.center = pt;
            }
            catch { }

            try
            {
                if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                {
                    data.radius_cm = ((Cylinder)face.Geometry).Radius;
                }
            }
            catch { }

            try
            {
                if (face.CreatedByFeature != null) data.created_by_feature = face.CreatedByFeature.Name;
            }
            catch { }

            return data;
        }

        static EdgeData ExtractEdge(Edge edge, ReferenceKeyManager mgr, int keyContext)
        {
            var data = new EdgeData
            {
                transient_key = edge.TransientKey,
                curve_type = edge.CurveType.ToString(),
                length_cm = GetEdgeLength(edge)
            };

            try
            {
                byte[] key = new byte[1];
                edge.GetReferenceKey(ref key, keyContext);
                data.reference_key_string = mgr.KeyToString(key);
            }
            catch { }

            try
            {
                CurveEvaluator eval = edge.Evaluator;
                double s, e;
                eval.GetParamExtents(out s, out e);
                double mid = (s + e) / 2.0;
                double[] arr = { mid };
                double[] pt = new double[3];
                double[] tan = new double[3];
                eval.GetPointAtParam(ref arr, ref pt);
                eval.GetTangent(ref arr, ref tan);
                data.midpoint = pt;
                data.tangent = tan;
                
                data.start_vertex = new double[] { edge.StartVertex.Point.X, edge.StartVertex.Point.Y, edge.StartVertex.Point.Z };
                data.end_vertex = new double[] { edge.StopVertex.Point.X, edge.StopVertex.Point.Y, edge.StopVertex.Point.Z };
            }
            catch { }

            // Extract Reference Keys of Adjacent Faces (No Hashing!)
            data.adjacent_faces = new List<string>();
            foreach (Face adjacentFace in edge.Faces)
            {
                try
                {
                    byte[] refKey = new byte[1];
                    adjacentFace.GetReferenceKey(ref refKey, keyContext);
                    data.adjacent_faces.Add(mgr.KeyToString(refKey));
                }
                catch { }
            }

            return data;
        }

        static double GetEdgeLength(Edge edge)
        {
            try
            {
                if (edge.CurveType == CurveTypeEnum.kLineCurve)
                {
                    Point p1 = edge.StartVertex.Point, p2 = edge.StopVertex.Point;
                    double dx = p2.X - p1.X, dy = p2.Y - p1.Y, dz = p2.Z - p1.Z;
                    return Math.Sqrt(dx * dx + dy * dy + dz * dz);
                }
                double st, en;
                edge.Evaluator.GetParamExtents(out st, out en);
                return Math.Abs(en - st);
            }
            catch { return 0; }
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
        public List<FaceData> faces { get; set; } = new List<FaceData>();
        public List<EdgeData> edges { get; set; } = new List<EdgeData>();
        public List<WorkPlaneData> work_planes { get; set; } = new List<WorkPlaneData>();
    }

    public class PartMetadata { public string file_name { get; set; } public string full_path { get; set; } public string internal_name { get; set; } public string units { get; set; } public string part_number { get; set; } public string description { get; set; } public string material { get; set; } public double mass_kg { get; set; } }
    public class BoundingBox { public PointData min { get; set; } public PointData max { get; set; } }
    public class PointData { public double x { get; set; } public double y { get; set; } public double z { get; set; } }

    public class WorkPlaneData
    {
        public string entity_type { get; set; } = "WorkPlane";
        public string work_feature_name { get; set; }
        public string surface_type { get; set; } = "kPlaneSurface";
        public double area_cm2 { get; set; } = 0.0001; 
        public double[] center { get; set; }
        public double[] normal { get; set; }
        public string reference_key_string { get; set; }
    }

    public class FaceData
    {
        public int transient_key { get; set; }
        public string reference_key_string { get; set; }
        public string surface_type { get; set; }
        public double area_cm2 { get; set; }
        public double[] normal { get; set; }
        public double[] center { get; set; }
        public double? radius_cm { get; set; }
        public PointData bbox_min { get; set; }
        public PointData bbox_max { get; set; }
        public string created_by_feature { get; set; }
    }

    public class EdgeData
    {
        public int transient_key { get; set; }
        public string reference_key_string { get; set; }
        public string curve_type { get; set; }
        public double length_cm { get; set; }
        public double[] midpoint { get; set; }
        public double[] tangent { get; set; }
        public double[] start_vertex { get; set; }
        public double[] end_vertex { get; set; }
        public List<string> adjacent_faces { get; set; }
    }
}