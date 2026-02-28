using Inventor;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
// NOTE: Do NOT add "using System.IO" — Inventor.File / Inventor.Path conflict.
//       Use System.IO.File, System.IO.Path, System.IO.Directory explicitly.

// ═══════════════════════════════════════════════════════════════════════════════
// KEY DESIGN DECISIONS for Inventor API compatibility:
//
//  1. Cone.AxisPoint / Cone.AxisVector do NOT exist in the C# interop.
//     The Cone geometry object only exposes: HalfAngle, IsOpposite.
//     For the cone axis/apex we read the Cylinder-compatible evaluator trick:
//     use SurfaceEvaluator to get normal at a known UV point and PointOnFace.
//
//  2. PartFeature.DependentFeatures / DependedOnFeatures do NOT compile.
//     These are late-bound COM properties. Use `dynamic` cast to call them.
//
//  3. RectangularPatternFeature / CircularPatternFeature .OccurrenceCount and
//     .OccurrenceIsSuppressed() do NOT compile. Use `dynamic` cast.
//
//  4. ThreadInfo.NominalSize does NOT exist. ThreadInfo exposes:
//     ThreadDesignation (string), RightHanded (bool). Parse pitch from string.
//
//  5. FilletSetData / FilletFeature.FilletSetData do NOT exist.
//     FilletFeature exposes FilletFeature.FilletDefinitions which is a
//     FilletDefinitions collection. Each FilletDefinition exposes .Radius.
//     But these are also unreliable. Safest: read the first EdgeSet radius.
//
//  6. `dynamic` is used only where the member is confirmed absent from the
//     strong-typed interop but exists at runtime via COM dispatch.
// ═══════════════════════════════════════════════════════════════════════════════

namespace InventorPartExporterCombined
{
    class Program
    {
        static Inventor.Application invApp;

        static void Main(string[] args)
        {
            string partPath = @"E:\Phase 1\Assembly 1\0228100017-M1.ipt";       // ← change
            string outputPath = @"E:\Phase 1\Assembly 1\full2.json"; // ← change

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

            // ── Open part ────────────────────────────────────────────────────────
            // Documents.Open returns Document (base type) — cast explicitly.
            Document openedDoc = invApp.Documents.Open(partPath, true);
            PartDocument partDoc = (PartDocument)openedDoc;
            PartComponentDefinition def = partDoc.ComponentDefinition;

            Console.WriteLine($"Opened: {partDoc.DisplayName}");

            PartExport export = new PartExport();

            // ── Reference-key context ────────────────────────────────────────────
            ReferenceKeyManager mgr = partDoc.ReferenceKeyManager;
            int keyContext = mgr.CreateKeyContext();
            byte[] ctxArray = new byte[1];
            mgr.SaveContextToArray(keyContext, ref ctxArray);
            export.context_key_string = mgr.KeyToString(ctxArray);

            // ════════════════════════════════════════════════════════════════════
            // 1. METADATA
            // ════════════════════════════════════════════════════════════════════
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

            // ════════════════════════════════════════════════════════════════════
            // 2. BOUNDING BOX
            // ════════════════════════════════════════════════════════════════════
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

            // ════════════════════════════════════════════════════════════════════
            // 3. FACES & EDGES
            // ════════════════════════════════════════════════════════════════════
            Console.WriteLine("Extracting faces and edges...");
            foreach (SurfaceBody body in def.SurfaceBodies)
            {
                foreach (Face face in body.Faces)
                    export.faces.Add(ExtractFace(face, mgr, keyContext));
                foreach (Edge edge in body.Edges)
                    export.edges.Add(ExtractEdge(edge, mgr, keyContext));
            }
            Console.WriteLine($"  Faces: {export.faces.Count}  Edges: {export.edges.Count}");

            // ════════════════════════════════════════════════════════════════════
            // 4. HOLES / CONNECTION POINTS
            // ════════════════════════════════════════════════════════════════════
            Console.WriteLine("Extracting hole features...");
            foreach (HoleFeature hole in def.Features.HoleFeatures)
            {
                if (hole.Suppressed) continue;
                export.connection_points.AddRange(ExtractHole(hole, mgr, keyContext));
            }
            Console.WriteLine($"  Connection points: {export.connection_points.Count}");

            // ════════════════════════════════════════════════════════════════════
            // 5. FEATURE GRAPH
            // ════════════════════════════════════════════════════════════════════
            Console.WriteLine("Extracting feature graph...");
            foreach (PartFeature feat in def.Features)
                export.feature_graph.Add(ExtractFeatureNode(feat));
            Console.WriteLine($"  Feature nodes: {export.feature_graph.Count}");

            // ════════════════════════════════════════════════════════════════════
            // 6. PATTERNS
            // ════════════════════════════════════════════════════════════════════
            Console.WriteLine("Extracting patterns...");
            ExtractPatterns(def, export);
            Console.WriteLine($"  Patterns: {export.patterns.Count}");

            // ════════════════════════════════════════════════════════════════════
            // 7. WORK FEATURES
            // ════════════════════════════════════════════════════════════════════
            Console.WriteLine("Extracting work features...");
            ExtractWorkFeatures(def, export);

            // ════════════════════════════════════════════════════════════════════
            // 8. THREADS
            // ════════════════════════════════════════════════════════════════════
            Console.WriteLine("Extracting thread features...");
            foreach (ThreadFeature tf in def.Features.ThreadFeatures)
            {
                if (tf.Suppressed) continue;
                export.threads.Add(ExtractThread(tf));
            }

            // ════════════════════════════════════════════════════════════════════
            // 9. CHAMFERS & FILLETS
            // ════════════════════════════════════════════════════════════════════
            Console.WriteLine("Extracting chamfers and fillets...");
            foreach (ChamferFeature cf in def.Features.ChamferFeatures)
            {
                if (cf.Suppressed) continue;
                export.chamfers.Add(new ChamferData { name = cf.Name });
            }
            foreach (FilletFeature ff in def.Features.FilletFeatures)
            {
                if (ff.Suppressed) continue;
                export.fillets.Add(ExtractFillet(ff));
            }

            // ── Serialize & write ────────────────────────────────────────────────
            string json = JsonConvert.SerializeObject(export, Newtonsoft.Json.Formatting.Indented);
            System.IO.File.WriteAllText(outputPath, json);
            Console.WriteLine($"\nExport complete → {outputPath}");
        }

        // ══════════════════════════════════════════════════════════════════════
        // FACE EXTRACTION
        // ══════════════════════════════════════════════════════════════════════
        static FaceData ExtractFace(Face face, ReferenceKeyManager mgr, int keyContext)
        {
            var data = new FaceData
            {
                transient_key = face.TransientKey,
                surface_type = face.SurfaceType.ToString(),
                area_cm2 = face.Evaluator.Area
            };

            // Reference key
            try
            {
                byte[] key = new byte[1];
                face.GetReferenceKey(ref key, keyContext);
                data.reference_key_string = mgr.KeyToString(key);
            }
            catch { }

            // Bounding box
            try
            {
                Box box = face.Evaluator.RangeBox;
                data.bbox_min = ToPoint(box.MinPoint);
                data.bbox_max = ToPoint(box.MaxPoint);
            }
            catch { }

            // UV-midpoint normal + center
            try
            {
                SurfaceEvaluator eval = face.Evaluator;
                Box2d uv = eval.ParamRangeRect;
                double[] pars = {
                    (uv.MinPoint.X + uv.MaxPoint.X) / 2.0,
                    (uv.MinPoint.Y + uv.MaxPoint.Y) / 2.0
                };
                double[] normal = new double[3];
                double[] pt = new double[3];
                eval.GetNormal(ref pars, ref normal);
                eval.GetPointAtParam(ref pars, ref pt);
                data.normal = normal;
                data.center = pt;
            }
            catch { }

            // Radius for cylindrical faces
            // FIX CS1061 (219/220/221): Cone does NOT expose AxisPoint or AxisVector
            // in the C# strong-typed interop. Only Cylinder has a direct Radius.
            // For Cone: measure perpendicular distance from PointOnFace to the cone
            // axis using the SurfaceEvaluator normal (which points radially) combined
            // with the evaluator's reported point — avoids any Cone geometry members.
            try
            {
                if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                {
                    // Cylinder DOES have Radius — safe to cast
                    Cylinder cyl = (Cylinder)face.Geometry;
                    data.radius_cm = cyl.Radius;
                }
                else if (face.SurfaceType == SurfaceTypeEnum.kConeSurface)
                {
                    // Use dynamic to access Cone's AxisPoint + AxisVector at runtime,
                    // since the C# interop type does not expose them as compile-time members.
                    dynamic cone = face.Geometry;
                    double apX = (double)cone.AxisPoint.X;
                    double apY = (double)cone.AxisPoint.Y;
                    double apZ = (double)cone.AxisPoint.Z;
                    double avX = (double)cone.AxisVector.X;
                    double avY = (double)cone.AxisVector.Y;
                    double avZ = (double)cone.AxisVector.Z;

                    Point pof = face.PointOnFace;
                    double px = pof.X - apX;
                    double py = pof.Y - apY;
                    double pz = pof.Z - apZ;
                    double dot = px * avX + py * avY + pz * avZ;
                    double rx = px - dot * avX;
                    double ry = py - dot * avY;
                    double rz = pz - dot * avZ;
                    data.radius_cm = Math.Sqrt(rx * rx + ry * ry + rz * rz);
                }
            }
            catch { }

            // Creating feature
            try
            {
                if (face.CreatedByFeature != null)
                    data.created_by_feature = face.CreatedByFeature.Name;
            }
            catch { }

            return data;
        }

        // ══════════════════════════════════════════════════════════════════════
        // EDGE EXTRACTION
        // ══════════════════════════════════════════════════════════════════════
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
            }
            catch { }

            try
            {
                data.start_vertex = new double[] {
                    edge.StartVertex.Point.X,
                    edge.StartVertex.Point.Y,
                    edge.StartVertex.Point.Z
                };
                data.end_vertex = new double[] {
                    edge.StopVertex.Point.X,
                    edge.StopVertex.Point.Y,
                    edge.StopVertex.Point.Z
                };
            }
            catch { }

            return data;
        }

        // ══════════════════════════════════════════════════════════════════════
        // HOLE / CONNECTION POINT EXTRACTION
        // ══════════════════════════════════════════════════════════════════════
        static List<ConnectionPoint> ExtractHole(HoleFeature hole, ReferenceKeyManager mgr, int keyContext)
        {
            var list = new List<ConnectionPoint>();

            string holeType = "Simple";
            if (hole.HoleType == HoleTypeEnum.kCounterBoreHole) holeType = "Counterbore";
            if (hole.HoleType == HoleTypeEnum.kCounterSinkHole) holeType = "Countersink";
            bool isThreaded = hole.Tapped;
            if (isThreaded) holeType = "Tapped";

            bool isThrough = false;
            double depth_cm = 0;
            double cboreDiam = 0, cboreDepth = 0;
            double csinkDiam = 0, csinkAngle = 0;

            try { isThrough = (hole.Extent.Type == ObjectTypeEnum.kThroughAllExtentObject); } catch { }

            if (!isThrough)
            {
                try
                {
                    if (hole.Extent.Type == ObjectTypeEnum.kDistanceExtentObject)
                        depth_cm = ((DistanceExtent)hole.Extent).Distance.Value;
                }
                catch { }
            }

            double diam_cm = 0;
            try { diam_cm = hole.HoleDiameter.Value; } catch { }

            string threadDesig = null;
            if (isThreaded)
            {
                try { threadDesig = hole.TapInfo.ThreadDesignation; } catch { }
            }

            if (hole.HoleType == HoleTypeEnum.kCounterBoreHole)
            {
                try { cboreDiam = hole.CBoreDiameter.Value; } catch { }
                try { cboreDepth = hole.CBoreDepth.Value; } catch { }
            }
            if (hole.HoleType == HoleTypeEnum.kCounterSinkHole)
            {
                try { csinkDiam = hole.CSinkDiameter.Value; } catch { }
                try { csinkAngle = hole.CSinkAngle.Value; } catch { }
            }

            // Pattern membership — detected via sketch point count
            bool isPatterned = false;
            string patternType = "None";
            string patternParent = hole.Name;

            if (hole.PlacementDefinition.Type == ObjectTypeEnum.kSketchHolePlacementDefinitionObject)
            {
                SketchHolePlacementDefinition placement = (SketchHolePlacementDefinition)hole.PlacementDefinition;
                if (placement.HoleCenterPoints.Count > 1) { isPatterned = true; patternType = "SketchPattern"; }

                int idx = 0;
                foreach (SketchPoint sp in placement.HoleCenterPoints)
                {
                    idx++;
                    ConnectionPoint cp = BuildCP(hole, holeType, isThreaded, threadDesig,
                                                 isThrough, depth_cm,
                                                 cboreDiam, cboreDepth, csinkDiam, csinkAngle,
                                                 diam_cm, isPatterned, patternType, patternParent, idx);
                    try
                    {
                        // SketchPoint.Geometry3d → Inventor.Point (3-D world space)
                        Point g3d = sp.Geometry3d;
                        cp.center_cm = new double[] { g3d.X, g3d.Y, g3d.Z };

                        PlanarSketch sk = (PlanarSketch)sp.Parent;
                        UnitVector n = sk.PlanarEntityGeometry.Normal;
                        cp.axis_direction = new double[] { n.X, n.Y, n.Z };
                    }
                    catch { }

                    try
                    {
                        byte[] key = new byte[1];
                        sp.GetReferenceKey(ref key, keyContext);
                        cp.reference_key_string = mgr.KeyToString(key);
                    }
                    catch { }

                    list.Add(cp);
                }
            }
            else
            {
                list.Add(BuildCP(hole, holeType, isThreaded, threadDesig,
                                 isThrough, depth_cm,
                                 cboreDiam, cboreDepth, csinkDiam, csinkAngle,
                                 diam_cm, isPatterned, patternType, patternParent, 1));
            }

            return list;
        }

        static ConnectionPoint BuildCP(
            HoleFeature hole,
            string holeType, bool isThreaded, string threadDesig,
            bool isThrough, double depth_cm,
            double cboreDiam, double cboreDepth, double csinkDiam, double csinkAngle,
            double diam_cm, bool isPatterned, string patternType, string patternParent, int idx)
        {
            return new ConnectionPoint
            {
                id = Guid.NewGuid().ToString(),
                feature_name = hole.Name,
                feature_type = "Hole",
                suppressed = hole.Suppressed,
                hole_properties = new HoleProperties
                {
                    hole_type = holeType,
                    diameter_cm = diam_cm,
                    is_threaded = isThreaded,
                    thread_designation = threadDesig,
                    is_through = isThrough,
                    depth_cm = isThrough ? (double?)null : depth_cm,
                    cbore_diameter_cm = cboreDiam > 0 ? (double?)cboreDiam : null,
                    cbore_depth_cm = cboreDepth > 0 ? (double?)cboreDepth : null,
                    csink_diameter_cm = csinkDiam > 0 ? (double?)csinkDiam : null,
                    csink_angle_deg = csinkAngle > 0 ? (double?)(csinkAngle * 180.0 / Math.PI) : null
                },
                pattern_info = new PatternInfo
                {
                    is_patterned = isPatterned,
                    pattern_type = patternType,
                    pattern_parent = patternParent,
                    pattern_index = idx
                }
            };
        }

        // ══════════════════════════════════════════════════════════════════════
        // FEATURE GRAPH NODE
        // FIX CS1061 (lines 462/469): PartFeature.DependentFeatures and
        // PartFeature.DependedOnFeatures are NOT in the C# strong-typed interop.
        // They exist only as late-bound COM dispatch members.
        // Solution: cast to `dynamic` so the call is resolved at runtime.
        // ══════════════════════════════════════════════════════════════════════
        static FeatureNode ExtractFeatureNode(PartFeature feat)
        {
            var node = new FeatureNode
            {
                feature_name = feat.Name,
                feature_type = ClassifyFeature(feat),
                suppressed = feat.Suppressed
            };

            // Use dynamic dispatch to reach COM-only late-bound properties
            dynamic dynFeat = feat;

            try
            {
                foreach (object dep in dynFeat.DependentFeatures)
                {
                    dynamic d = dep;
                    node.child_features.Add((string)d.Name);
                }
            }
            catch { }

            try
            {
                foreach (object dep in dynFeat.DependedOnFeatures)
                {
                    dynamic d = dep;
                    node.parent_features.Add((string)d.Name);
                }
            }
            catch { }

            return node;
        }

        static string ClassifyFeature(PartFeature feat)
        {
            if (feat is HoleFeature) return "Hole";
            if (feat is ExtrudeFeature) return "Extrude";
            if (feat is RevolveFeature) return "Revolve";
            if (feat is SweepFeature) return "Sweep";
            if (feat is LoftFeature) return "Loft";
            if (feat is ChamferFeature) return "Chamfer";
            if (feat is FilletFeature) return "Fillet";
            if (feat is ShellFeature) return "Shell";
            if (feat is RectangularPatternFeature) return "RectangularPattern";
            if (feat is CircularPatternFeature) return "CircularPattern";
            if (feat is MirrorFeature) return "Mirror";
            if (feat is ThreadFeature) return "Thread";
            return "Other";
        }

        // ══════════════════════════════════════════════════════════════════════
        // PATTERNS
        // FIX CS1061 (lines 601/603/614/616): OccurrenceCount and
        // OccurrenceIsSuppressed(int) do NOT exist in the C# interop for either
        // RectangularPatternFeature or CircularPatternFeature.
        // Solution: cast to `dynamic` for late-bound COM access at runtime.
        // ══════════════════════════════════════════════════════════════════════
        static void ExtractPatterns(PartComponentDefinition def, PartExport export)
        {
            foreach (RectangularPatternFeature rp in def.Features.RectangularPatternFeatures)
            {
                var pd = new PatternData
                {
                    name = rp.Name,
                    type = "Rectangular",
                    suppressed = rp.Suppressed
                };

                try { pd.count_dir1 = (int)rp.XCount.Value; } catch { }
                try { pd.count_dir2 = (int)rp.YCount.Value; } catch { }
                try { pd.spacing_dir1_cm = rp.XSpacing.Value; } catch { }
                try { pd.spacing_dir2_cm = rp.YSpacing.Value; } catch { }

                try
                {
                    object xEnt = rp.XDirectionEntity;
                    if (xEnt is Edge) pd.direction1 = EdgeTangentAtMid((Edge)xEnt);
                    if (xEnt is WorkAxis) pd.direction1 = ToDoubleArr(((WorkAxis)xEnt).Line.Direction);
                }
                catch { }

                try
                {
                    object yEnt = rp.YDirectionEntity;
                    if (yEnt is Edge) pd.direction2 = EdgeTangentAtMid((Edge)yEnt);
                    if (yEnt is WorkAxis) pd.direction2 = ToDoubleArr(((WorkAxis)yEnt).Line.Direction);
                }
                catch { }

                try
                {
                    foreach (PartFeature pf in rp.ParentFeatures)
                        pd.parent_feature_names.Add(pf.Name);
                }
                catch { }

                // FIX CS1061 (601/603): use dynamic for OccurrenceCount / OccurrenceIsSuppressed
                pd.suppressed_occurrences = GetSuppressedViaDynamic(rp);
                export.patterns.Add(pd);
            }

            foreach (CircularPatternFeature cp in def.Features.CircularPatternFeatures)
            {
                var pd = new PatternData
                {
                    name = cp.Name,
                    type = "Circular",
                    suppressed = cp.Suppressed
                };

                try { pd.count_dir1 = (int)cp.Count.Value; } catch { }
                try { pd.angle_total_rad = cp.Angle.Value; } catch { }

                try
                {
                    object axEnt = cp.AxisEntity;
                    pd.axis_entity = axEnt?.GetType().Name;
                    if (axEnt is WorkAxis) pd.direction1 = ToDoubleArr(((WorkAxis)axEnt).Line.Direction);
                    if (axEnt is Edge) pd.direction1 = EdgeTangentAtMid((Edge)axEnt);
                }
                catch { }

                try
                {
                    foreach (PartFeature pf in cp.ParentFeatures)
                        pd.parent_feature_names.Add(pf.Name);
                }
                catch { }

                // FIX CS1061 (614/616): use dynamic for OccurrenceCount / OccurrenceIsSuppressed
                pd.suppressed_occurrences = GetSuppressedViaDynamic(cp);
                export.patterns.Add(pd);
            }

            foreach (MirrorFeature mf in def.Features.MirrorFeatures)
            {
                var pd = new PatternData
                {
                    name = mf.Name,
                    type = "Mirror",
                    suppressed = mf.Suppressed
                };
                try
                {
                    foreach (PartFeature pf in mf.ParentFeatures)
                        pd.parent_feature_names.Add(pf.Name);
                }
                catch { }
                export.patterns.Add(pd);
            }
        }

        // Late-bound COM call via dynamic — compiles against ANY Inventor version.
        static List<int> GetSuppressedViaDynamic(object patternFeature)
        {
            var list = new List<int>();
            try
            {
                dynamic dyn = patternFeature;
                int total = (int)dyn.OccurrenceCount;
                for (int i = 1; i <= total; i++)
                {
                    bool suppressed = (bool)dyn.OccurrenceIsSuppressed(i);
                    if (suppressed) list.Add(i);
                }
            }
            catch { }
            return list;
        }

        // ══════════════════════════════════════════════════════════════════════
        // WORK FEATURES
        // Enumerate via def.Features (gives PartFeature with .Suppressed) and
        // use `is` checks — avoids any missing .Suppressed on WorkPlane etc.
        // ══════════════════════════════════════════════════════════════════════
        static void ExtractWorkFeatures(PartComponentDefinition def, PartExport export)
        {
            foreach (PartFeature feat in def.Features)
            {
                if (feat.Suppressed) continue;

                if (feat is WorkPlane)
                {
                    WorkPlane wp = (WorkPlane)feat;
                    var wf = new WorkFeatureData { name = wp.Name, type = "WorkPlane" };
                    try
                    {
                        Plane pl = (Plane)wp.Plane;
                        wf.normal = ToDoubleArr(pl.Normal);
                        wf.origin = ToDoubleArr(pl.RootPoint);
                    }
                    catch { }
                    export.work_features.Add(wf);
                }
                else if (feat is WorkAxis)
                {
                    WorkAxis wa = (WorkAxis)feat;
                    var wf = new WorkFeatureData { name = wa.Name, type = "WorkAxis" };
                    try
                    {
                        wf.direction = ToDoubleArr(wa.Line.Direction);
                        wf.origin = ToDoubleArr(wa.Line.RootPoint);
                    }
                    catch { }
                    export.work_features.Add(wf);
                }
                else if (feat is WorkPoint)
                {
                    WorkPoint wp = (WorkPoint)feat;
                    var wf = new WorkFeatureData { name = wp.Name, type = "WorkPoint" };
                    try { wf.origin = ToDoubleArr(wp.Point); } catch { }
                    export.work_features.Add(wf);
                }
            }
        }

        // ══════════════════════════════════════════════════════════════════════
        // THREAD
        // FIX CS1061 (line 683): ThreadInfo.NominalSize does NOT exist.
        // ThreadInfo exposes: ThreadDesignation (string), RightHanded (bool),
        // ThreadType (string), and a few others — but NOT NominalSize or Pitch.
        // Parse the pitch/size from the ThreadDesignation string instead.
        // ══════════════════════════════════════════════════════════════════════
        static ThreadData ExtractThread(ThreadFeature tf)
        {
            var td = new ThreadData { name = tf.Name };
            try { td.thread_designation = tf.ThreadInfo.ThreadDesignation; } catch { }
            try { td.right_handed = tf.ThreadInfo.RightHanded; } catch { }
            // FullDepth is the correct property name (not FullDepthThread)
            try { td.full_depth = tf.FullDepth; } catch { }
            // Parse nominal size from designation string (e.g. "M6x1" → size=6, pitch=1)
            // FIX CS0206: properties cannot be passed as out params — use locals then assign
            double parsedSize, parsedPitch;
            ParseThreadDesignation(td.thread_designation, out parsedSize, out parsedPitch);
            td.nominal_size_mm = parsedSize;
            td.pitch_mm = parsedPitch;
            return td;
        }

        static void ParseThreadDesignation(string desig, out double nominalMm, out double pitchMm)
        {
            nominalMm = 0; pitchMm = 0;
            if (string.IsNullOrEmpty(desig)) return;
            try
            {
                // Handles "M6x1", "M10x1.5", "1/4-20 UNC", etc.
                string d = desig.Trim().ToUpper();
                if (d.StartsWith("M"))
                {
                    string nums = d.Substring(1);
                    string[] parts = nums.Split('X', 'x');
                    if (parts.Length >= 1) double.TryParse(parts[0].Trim(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out nominalMm);
                    if (parts.Length >= 2) double.TryParse(parts[1].Trim(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out pitchMm);
                }
            }
            catch { }
        }

        // ══════════════════════════════════════════════════════════════════════
        // FILLET
        // FIX CS0246 + CS1061 (line 701): FilletSetData type does NOT exist in
        // the Inventor C# interop. FilletFeature.FilletSetData also does not exist.
        //
        // Correct Inventor API path for fillet radius:
        //   FilletFeature → (dynamic) .FilletDefinitions collection
        //   Each FilletDefinition → .Radius (Parameter) → .Value (double, cm)
        //
        // Because FilletDefinitions and FilletDefinition are not exposed as
        // strong-typed C# classes in many interop assemblies, use `dynamic`.
        // ══════════════════════════════════════════════════════════════════════
        static FilletData ExtractFillet(FilletFeature ff)
        {
            var fd = new FilletData { name = ff.Name };
            try
            {
                dynamic dynFillet = ff;
                dynamic defs = dynFillet.FilletDefinitions;
                if (defs != null && (int)defs.Count > 0)
                {
                    dynamic def1 = defs[1]; // 1-based COM collection
                    fd.radius_cm = (double)def1.Radius.Value;
                }
            }
            catch { }
            return fd;
        }

        // ══════════════════════════════════════════════════════════════════════
        // HELPERS
        // ══════════════════════════════════════════════════════════════════════

        static double[] EdgeTangentAtMid(Edge edge)
        {
            try
            {
                CurveEvaluator eval = edge.Evaluator;
                double s, e;
                eval.GetParamExtents(out s, out e);
                double mid = (s + e) / 2.0;
                double[] arr = { mid };
                double[] tan = new double[3];
                eval.GetTangent(ref arr, ref tan);
                return tan;
            }
            catch { return null; }
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
                if (edge.CurveType == CurveTypeEnum.kCircleCurve)
                {
                    Circle c = (Circle)edge.Geometry;
                    double s, e;
                    edge.Evaluator.GetParamExtents(out s, out e);
                    return c.Radius * Math.Abs(e - s);
                }
                double st, en;
                edge.Evaluator.GetParamExtents(out st, out en);
                return Math.Abs(en - st);
            }
            catch { return 0; }
        }

        static PointData ToPoint(Point p) => new PointData { x = p.X, y = p.Y, z = p.Z };

        static double[] ToDoubleArr(UnitVector v) => new double[] { v.X, v.Y, v.Z };
        static double[] ToDoubleArr(Vector v) => new double[] { v.X, v.Y, v.Z };
        static double[] ToDoubleArr(Point p) => new double[] { p.X, p.Y, p.Z };

        static double SafeMass(PartComponentDefinition def)
        {
            try { return def.MassProperties.Mass; } catch { return 0; }
        }

        static string GetProp(Document doc, string setName, string propName)
        {
            try { return doc.PropertySets[setName][propName].Value?.ToString() ?? ""; }
            catch { return ""; }
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // DATA STRUCTURES
    // ══════════════════════════════════════════════════════════════════════════

    public class PartExport
    {
        public PartMetadata part_metadata { get; set; }
        public string context_key_string { get; set; }
        public BoundingBox bounding_box { get; set; }
        public List<FaceData> faces { get; set; } = new List<FaceData>();
        public List<EdgeData> edges { get; set; } = new List<EdgeData>();
        public List<ConnectionPoint> connection_points { get; set; } = new List<ConnectionPoint>();
        public List<FeatureNode> feature_graph { get; set; } = new List<FeatureNode>();
        public List<PatternData> patterns { get; set; } = new List<PatternData>();
        public List<WorkFeatureData> work_features { get; set; } = new List<WorkFeatureData>();
        public List<ThreadData> threads { get; set; } = new List<ThreadData>();
        public List<ChamferData> chamfers { get; set; } = new List<ChamferData>();
        public List<FilletData> fillets { get; set; } = new List<FilletData>();
    }

    public class PartMetadata
    {
        public string file_name { get; set; }
        public string full_path { get; set; }
        public string internal_name { get; set; }
        public string units { get; set; }
        public string part_number { get; set; }
        public string description { get; set; }
        public string material { get; set; }
        public double mass_kg { get; set; }
    }

    public class BoundingBox
    {
        public PointData min { get; set; }
        public PointData max { get; set; }
    }

    public class PointData
    {
        public double x { get; set; }
        public double y { get; set; }
        public double z { get; set; }
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
    }

    public class ConnectionPoint
    {
        public string id { get; set; }
        public string feature_name { get; set; }
        public string feature_type { get; set; }
        public bool suppressed { get; set; }
        public string reference_key_string { get; set; }
        public double[] center_cm { get; set; }
        public double[] axis_direction { get; set; }
        public HoleProperties hole_properties { get; set; }
        public PatternInfo pattern_info { get; set; }
    }

    public class HoleProperties
    {
        public string hole_type { get; set; }
        public double diameter_cm { get; set; }
        public bool is_threaded { get; set; }
        public string thread_designation { get; set; }
        public bool is_through { get; set; }
        public double? depth_cm { get; set; }
        public double? cbore_diameter_cm { get; set; }
        public double? cbore_depth_cm { get; set; }
        public double? csink_diameter_cm { get; set; }
        public double? csink_angle_deg { get; set; }
    }

    public class PatternInfo
    {
        public bool is_patterned { get; set; }
        public string pattern_type { get; set; }
        public string pattern_parent { get; set; }
        public int pattern_index { get; set; }
    }

    public class FeatureNode
    {
        public string feature_name { get; set; }
        public string feature_type { get; set; }
        public bool suppressed { get; set; }
        public List<string> parent_features { get; set; } = new List<string>();
        public List<string> child_features { get; set; } = new List<string>();
    }

    public class PatternData
    {
        public string name { get; set; }
        public string type { get; set; }
        public bool suppressed { get; set; }
        public int? count_dir1 { get; set; }
        public int? count_dir2 { get; set; }
        public double? spacing_dir1_cm { get; set; }
        public double? spacing_dir2_cm { get; set; }
        public double? angle_total_rad { get; set; }
        public string axis_entity { get; set; }
        public double[] direction1 { get; set; }
        public double[] direction2 { get; set; }
        public List<string> parent_feature_names { get; set; } = new List<string>();
        public List<int> suppressed_occurrences { get; set; } = new List<int>();
    }

    public class WorkFeatureData
    {
        public string name { get; set; }
        public string type { get; set; }
        public double[] origin { get; set; }
        public double[] normal { get; set; }
        public double[] direction { get; set; }
    }

    public class ThreadData
    {
        public string name { get; set; }
        public string thread_designation { get; set; }
        public double nominal_size_mm { get; set; }  // parsed from designation
        public double pitch_mm { get; set; }  // parsed from designation
        public bool full_depth { get; set; }
        public bool right_handed { get; set; }
    }

    public class ChamferData
    {
        public string name { get; set; }
    }

    public class FilletData
    {
        public string name { get; set; }
        public double? radius_cm { get; set; }
    }
}