using Inventor;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace InventorAssemblyExporter
{
    class Program
    {
        static Inventor.Application invApp;

        static Dictionary<string, KeyContextData> contextCache = new Dictionary<string, KeyContextData>();

        static void Main(string[] args)
        {
            string baseInputRoot = @"E:\Phase 1";
            string baseOutputRoot = @"E:\Phase 1\assembliesexport";

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

            Console.WriteLine($"Starting Generative Assembly AI extraction in: {baseInputRoot}");

            if (!System.IO.Directory.Exists(baseOutputRoot))
                System.IO.Directory.CreateDirectory(baseOutputRoot);

            foreach (string iamPath in System.IO.Directory.GetFiles(baseInputRoot, "*.iam", System.IO.SearchOption.AllDirectories))
            {
                ProcessAssembly(iamPath, baseOutputRoot);
            }

            Console.WriteLine("\n✅ All assembly graphs successfully extracted for GNN.");
        }

        static void ProcessAssembly(string iamPath, string outputRoot)
        {
            try
            {
                Console.WriteLine($"\nOpening Assembly: {iamPath}");
                contextCache.Clear(); // Clear cache for new assembly

                Document openedDoc = invApp.Documents.Open(iamPath, true);
                AssemblyDocument asmDoc = (AssemblyDocument)openedDoc;
                AssemblyComponentDefinition def = asmDoc.ComponentDefinition;

                AssemblyExport export = new AssemblyExport();

                export.assembly_metadata = new AssemblyMetadata
                {
                    assembly_name = asmDoc.DisplayName,
                    full_file_name = asmDoc.FullFileName,
                    internal_name = asmDoc.InternalName,
                    total_occurrences = def.Occurrences.Count,
                    total_constraints = def.Constraints.Count
                };

                // --- 1. ASSEMBLY PHYSICS ---
                try
                {
                    MassProperties massProps = def.MassProperties;
                    export.physics = new AssemblyPhysics
                    {
                        mass_kg = massProps.Mass,
                        center_of_mass = new double[] { massProps.CenterOfMass.X, massProps.CenterOfMass.Y, massProps.CenterOfMass.Z }
                    };
                    double Ixx, Iyy, Izz, Ixy, Iyz, Ixz;
                    massProps.XYZMomentsOfInertia(out Ixx, out Iyy, out Izz, out Ixy, out Iyz, out Ixz);
                    export.physics.inertia_tensor = new double[] { Ixx, Iyy, Izz, Ixy, Iyz, Ixz };
                } catch { }

                // --- 2. GRAPH NODES (Occurrences + DOF) ---
                Console.WriteLine("  Extracting Graph Nodes & DOFs...");
                ExtractOccurrences(def.Occurrences, export.assembly_graph.nodes, "");

                // --- 3. GRAPH EDGES (Constraints) ---
                Console.WriteLine("  Extracting Constraint Edges (with O(1) Cache)...");
                foreach (AssemblyConstraint constraint in def.Constraints)
                {
                    var edge = new ConstraintEdge
                    {
                        constraint_name = constraint.Name,
                        constraint_type = constraint.Type.ToString().Replace("Object", ""),
                        suppressed = constraint.Suppressed,
                        health_status = constraint.HealthStatus.ToString()
                    };

                    dynamic dynConst = constraint;
                    try { edge.node_one_id = constraint.OccurrenceOne.Name; } catch { }
                    try { edge.node_two_id = constraint.OccurrenceTwo.Name; } catch { }
                    try { edge.offset_cm = (double)dynConst.Offset.Value; } catch { }
                    try { edge.angle_rad = (double)dynConst.Angle.Value; } catch { }

                    try { edge.entity_one = ExtractEntity(constraint.EntityOne); } catch { }
                    try { edge.entity_two = ExtractEntity(constraint.EntityTwo); } catch { }

                    export.assembly_graph.constraint_edges.Add(edge);
                }

                // --- 4. GRAPH EDGES (Kinematic Joints) ---
                Console.WriteLine("  Extracting Kinematic Joint Edges...");
                try
                {
                    foreach (AssemblyJoint joint in def.Joints)
                    {
                        var jEdge = new JointEdge
                        {
                            joint_name = joint.Name,
                            joint_type = joint.Definition.JointType.ToString(),
                            health_status = joint.HealthStatus.ToString(),
                            suppressed = joint.Suppressed,
                            node_one_id = joint.OccurrenceOne?.Name,
                            node_two_id = joint.OccurrenceTwo?.Name,
                            has_linear_limit = joint.Definition.HasLinearPositionLimits,
                            has_angular_limit = joint.Definition.HasAngularPositionLimits
                        };
                        
                        if (jEdge.has_linear_limit) {
                            jEdge.linear_start_cm = joint.Definition.LinearPositionStartLimit.Value;
                            jEdge.linear_end_cm = joint.Definition.LinearPositionEndLimit.Value;
                        }
                        if (jEdge.has_angular_limit) {
                            jEdge.angular_start_rad = joint.Definition.AngularPositionStartLimit.Value;
                            jEdge.angular_end_rad = joint.Definition.AngularPositionEndLimit.Value;
                        }
                        
                        export.assembly_graph.joint_edges.Add(jEdge);
                    }
                } catch { }

                string fileName = System.IO.Path.GetFileNameWithoutExtension(iamPath);
                string outputPath = System.IO.Path.Combine(outputRoot, fileName + ".json");
                System.IO.File.WriteAllText(outputPath, JsonConvert.SerializeObject(export, Newtonsoft.Json.Formatting.Indented));

                asmDoc.Close(true);
                Console.WriteLine($"  Exported → {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Error] Failed processing {iamPath}: {ex.Message}");
            }
        }

        // ==========================================
        // GRAPH NODES (Occurrences + DOF Extraction)
        // ==========================================
        static void ExtractOccurrences(ComponentOccurrences occurrences, List<ComponentNode> nodes, string parentPath)
        {
            foreach (ComponentOccurrence occ in occurrences)
            {
                string nodeId = string.IsNullOrEmpty(parentPath) ? occ.Name : parentPath + "/" + occ.Name;

                ComponentNode node = new ComponentNode
                {
                    node_id = nodeId,
                    file_name = occ.ReferencedDocumentDescriptor?.FullDocumentName ?? "",
                    component_type = occ.DefinitionDocumentType.ToString(),
                    grounded = occ.Grounded,
                    suppressed = occ.Suppressed,
                    visible = occ.Visible,
                    adaptive = occ.Adaptive,
                    transform = ExtractTransform(occ.Transformation)
                };

                try { node.flexible = occ.Flexible; } catch { }

                // EXTRACT KINEMATIC DEGREES OF FREEDOM
                try
                {
                    int transCount = 0, rotCount = 0;
                    ObjectsEnumerator transDOFs, rotDOFs;
                    Point dofCenter;
                    occ.GetDegreesOfFreedom(out transCount, out transDOFs, out rotCount, out rotDOFs, out dofCenter);
                    
                    node.dof_translation = transCount;
                    node.dof_rotation = rotCount;
                    node.fully_constrained = (transCount == 0 && rotCount == 0 && !occ.Grounded);
                }
                catch { }

                nodes.Add(node);

                if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    try { ExtractOccurrences(occ.SubOccurrences, nodes, nodeId); } catch { }
                }
            }
        }

        // ==========================================
        // FAST REACH-THROUGH NATIVE EXTRACTION
        // ==========================================
        static EntityData ExtractEntity(object obj)
        {
            if (obj == null) return null;
            var data = new EntityData();

            try
            {
                dynamic proxy = obj;
                data.entity_type = proxy.Type.ToString().Replace("k", "").Replace("Object", "");

                ComponentOccurrence occ = proxy.ContainingOccurrence;
                data.proxy_context_occurrence = occ.Name;

                Document doc = (Document)occ.Definition.Document;
                string docPath = doc.FullFileName;
                data.owner_document = docPath;

                // --- HIGH PERFORMANCE CONTEXT CACHING ---
                KeyContextData ctxData;
                if (!contextCache.TryGetValue(docPath, out ctxData))
                {
                    ReferenceKeyManager mgr = doc.ReferenceKeyManager;
                    int ctxId = mgr.CreateKeyContext();
                    byte[] ctxBytes = new byte[0]; // CORRECT DYNAMIC ALLOCATION
                    mgr.SaveContextToArray(ctxId, ref ctxBytes);
                    
                    ctxData = new KeyContextData { ContextId = ctxId, ContextString = mgr.KeyToString(ctxBytes), Manager = mgr };
                    contextCache[docPath] = ctxData;
                }

                data.context_key_string = ctxData.ContextString;

                object nativeObj = proxy.NativeObject;
                dynamic nativeDyn = nativeObj;

                byte[] refKey = new byte[0]; // CORRECT DYNAMIC ALLOCATION
                nativeDyn.GetReferenceKey(ref refKey, ctxData.ContextId);
                data.reference_key_string = ctxData.Manager.KeyToString(refKey);

                // --- EXTRACT PHYSICAL MATH ---
                if (nativeObj is Face)
                {
                    Face face = (Face)nativeObj;
                    data.surface_type = face.SurfaceType.ToString();
                    data.area_cm2 = face.Evaluator.Area;

                    try
                    {
                        Box box = face.Evaluator.RangeBox;
                        data.face_bbox_min = new double[] { box.MinPoint.X, box.MinPoint.Y, box.MinPoint.Z };
                        data.face_bbox_max = new double[] { box.MaxPoint.X, box.MaxPoint.Y, box.MaxPoint.Z };
                    } catch { }

                    try
                    {
                        SurfaceEvaluator eval = face.Evaluator;
                        Box2d uv = eval.ParamRangeRect;
                        double[] pars = { (uv.MinPoint.X + uv.MaxPoint.X) / 2.0, (uv.MinPoint.Y + uv.MaxPoint.Y) / 2.0 };
                        double[] normal = new double[3], pt = new double[3];
                        eval.GetNormal(ref pars, ref normal);
                        eval.GetPointAtParam(ref pars, ref pt);
                        data.face_normal_at_center = normal;
                        data.face_point_at_center = pt;
                    } catch { }

                    data.loops = new List<LoopData>();
                    foreach (EdgeLoop loop in face.EdgeLoops)
                    {
                        var loopData = new LoopData { is_outer = loop.IsOuterEdgeLoop };
                        foreach (Edge loopEdge in loop.Edges)
                        {
                            try
                            {
                                byte[] eKey = new byte[0];
                                loopEdge.GetReferenceKey(ref eKey, ctxData.ContextId);
                                loopData.edge_reference_keys.Add(ctxData.Manager.KeyToString(eKey));
                            } catch { }
                        }
                        data.loops.Add(loopData);
                    }
                }
                else if (nativeObj is Edge)
                {
                    Edge edge = (Edge)nativeObj;
                    data.curve_type = edge.CurveType.ToString();

                    try
                    {
                        CurveEvaluator eval = edge.Evaluator;
                        double s, e;
                        eval.GetParamExtents(out s, out e);
                        
                        double length;
                        eval.GetLengthAtParam(s, e, out length);
                        data.length_cm = length;

                        double midParam;
                        eval.GetParamAtLength(s, length / 2.0, out midParam);

                        double[] arr = { midParam }, pt = new double[3], tan = new double[3];
                        eval.GetPointAtParam(ref arr, ref pt);
                        eval.GetTangent(ref arr, ref tan);
                        
                        data.edge_midpoint = pt;
                        data.edge_tangent_at_mid = tan;
                    } catch { }
                }
                else if (nativeObj is WorkPlane)
                {
                    WorkPlane wp = (WorkPlane)nativeObj;
                    data.work_feature_name = wp.Name;
                    data.surface_type = "kPlaneSurface";
                    data.area_cm2 = 0.0001; 
                    try
                    {
                        Plane pl = (Plane)wp.Plane;
                        data.face_normal_at_center = new double[] { pl.Normal.X, pl.Normal.Y, pl.Normal.Z };
                        data.face_point_at_center = new double[] { pl.RootPoint.X, pl.RootPoint.Y, pl.RootPoint.Z };
                    } catch { }
                }
            }
            catch { }

            return data;
        }

        static TransformData ExtractTransform(Matrix matrix)
        {
            return new TransformData
            {
                rotation_matrix = new double[][]
                {
                    new double[] { matrix.get_Cell(1, 1), matrix.get_Cell(1, 2), matrix.get_Cell(1, 3) },
                    new double[] { matrix.get_Cell(2, 1), matrix.get_Cell(2, 2), matrix.get_Cell(2, 3) },
                    new double[] { matrix.get_Cell(3, 1), matrix.get_Cell(3, 2), matrix.get_Cell(3, 3) }
                },
                translation_cm = new double[] { matrix.get_Cell(1, 4), matrix.get_Cell(2, 4), matrix.get_Cell(3, 4) }
            };
        }
    }

    // ==========================================
    // DATA STRUCTURES (GNN OPTIMIZED)
    // ==========================================
    public class KeyContextData
    {
        public int ContextId;
        public string ContextString;
        public ReferenceKeyManager Manager;
    }

    public class AssemblyExport
    {
        public AssemblyMetadata assembly_metadata { get; set; }
        public AssemblyPhysics physics { get; set; }
        public AssemblyGraph assembly_graph { get; set; } = new AssemblyGraph();
    }

    public class AssemblyMetadata { public string assembly_name { get; set; } public string full_file_name { get; set; } public string internal_name { get; set; } public int total_occurrences { get; set; } public int total_constraints { get; set; } }
    public class AssemblyPhysics { public double mass_kg { get; set; } public double[] center_of_mass { get; set; } public double[] inertia_tensor { get; set; } }
    
    // EXPLICIT GRAPH STRUCTURE FOR ML
    public class AssemblyGraph
    {
        public List<ComponentNode> nodes { get; set; } = new List<ComponentNode>();
        public List<ConstraintEdge> constraint_edges { get; set; } = new List<ConstraintEdge>();
        public List<JointEdge> joint_edges { get; set; } = new List<JointEdge>();
    }

    public class ComponentNode
    {
        public string node_id { get; set; } // The full occurrence path
        public string file_name { get; set; }
        public string component_type { get; set; }
        public bool grounded { get; set; }
        public bool suppressed { get; set; }
        public bool visible { get; set; }
        public bool adaptive { get; set; }
        public bool flexible { get; set; }
        public int dof_translation { get; set; }
        public int dof_rotation { get; set; }
        public bool fully_constrained { get; set; }
        public TransformData transform { get; set; }
    }

    public class TransformData { public double[][] rotation_matrix { get; set; } public double[] translation_cm { get; set; } }

    public class ConstraintEdge
    {
        public string constraint_name { get; set; }
        public string constraint_type { get; set; }
        public bool suppressed { get; set; }
        public string health_status { get; set; } 
        public string node_one_id { get; set; }
        public string node_two_id { get; set; }
        public double? offset_cm { get; set; }
        public double? angle_rad { get; set; }
        public EntityData entity_one { get; set; }
        public EntityData entity_two { get; set; }
    }

    public class JointEdge
    {
        public string joint_name { get; set; }
        public string joint_type { get; set; }
        public string health_status { get; set; }
        public bool suppressed { get; set; }
        public string node_one_id { get; set; }
        public string node_two_id { get; set; }
        public bool has_linear_limit { get; set; }
        public double? linear_start_cm { get; set; }
        public double? linear_end_cm { get; set; }
        public bool has_angular_limit { get; set; }
        public double? angular_start_rad { get; set; }
        public double? angular_end_rad { get; set; }
    }

    public class EntityData
    {
        public string entity_type { get; set; }
        public string proxy_context_occurrence { get; set; }
        public string owner_document { get; set; }
        public string reference_key_string { get; set; }
        public string context_key_string { get; set; }
        
        public string surface_type { get; set; }
        public double? area_cm2 { get; set; }
        public double[] face_normal_at_center { get; set; }
        public double[] face_point_at_center { get; set; }
        public double[] face_bbox_min { get; set; }
        public double[] face_bbox_max { get; set; }
        public List<LoopData> loops { get; set; } = new List<LoopData>();
        
        public string curve_type { get; set; }
        public double? length_cm { get; set; }
        public double[] edge_midpoint { get; set; }
        public double[] edge_tangent_at_mid { get; set; }
        
        public string work_feature_name { get; set; }
    }

    public class LoopData { public bool is_outer { get; set; } public List<string> edge_reference_keys { get; set; } = new List<string>(); }
}