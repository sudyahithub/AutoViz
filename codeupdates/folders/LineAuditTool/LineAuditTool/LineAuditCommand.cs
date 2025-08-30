// LineAuditCommand.cs  (C# 7.3 compatible)
//
// Command: MATCHANDREPLACECHAIRS
// - Uses the new Matching layer (Preview + pHash/dHash + Geometry + Scoring) to rank candidates.
// - Keeps your WinForms preview picker, pivot alignment and NCC-based rotation correction.
// - Falls back to your previous visual/name ranking if matcher init is unavailable.
//
// Build x64 against your AutoCAD version.

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using DrawingColor = System.Drawing.Color;

[assembly: CommandClass(typeof(LineAuditTool.BlockMatcher))]

namespace LineAuditTool
{
    public class BlockMatcher
    {
        // ======== CONFIG: paths ========
        private const string MasterDwgPath = @"C:\Users\admin\Downloads\VIZ-AUTOCAD\M1.dwg";

        private const string PythonExportsRoot = @"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS";
        private const string InstancesDirName = "INSTANCES";
        private const string AutoClustersDirName = "AUTO_CLUSTERS";

        private const string PythonExe = "python";
        private const string PythonExporterScript = @"C:\Users\admin\Downloads\VIZ-AUTOCAD\SCRIPTS\one.py";
        private const int PythonTimeoutMs = 25000;
        private static bool _exporterRunThisSession = false;

        // ======== (legacy) visual prefilter weights (fallback path only) ========
        private const int NamePrefilterTopK = 60;
        private const double W_VIS = 0.50;   // dHash
        private const double W_SHAPE = 0.30; // radial fingerprint
        private const double W_SIZE = 0.15;  // area/aspect
        private const double W_NAME = 0.05;  // name tie-breaker
        private const double SIZE_GATE = 0.25;
        private const int SHAPE_RAYS = 48;
        private const byte FG_THRESH = 230;

        // ======== NEW: Matching singletons (init once per session) ========
        private static LineAuditTool.Matching.MasterIndex _masterIdx;
        private static LineAuditTool.Matching.MatchConfig _matchCfg;
        private static LineAuditTool.Matching.PreviewRenderer _preview;

        private static LineAuditTool.Matching.MatchConfig LoadConfigOrDefault()
        {
            var cfg = new LineAuditTool.Matching.MatchConfig();
            try
            {
                string env = Environment.GetEnvironmentVariable("BLOCKMATCH_CFG");
                if (!string.IsNullOrEmpty(env) && File.Exists(env))
                {
                    string json = File.ReadAllText(env);
                    var loaded = Newtonsoft.Json.JsonConvert.DeserializeObject<LineAuditTool.Matching.MatchConfig>(json);
                    if (loaded != null) cfg = loaded;
                }
                else
                {
                    // Look next to the plugin DLL: <AppBase>\Config\MatchConfig.json
                    string local = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "MatchConfig.json");
                    if (File.Exists(local))
                    {
                        string json = File.ReadAllText(local);
                        var loaded = Newtonsoft.Json.JsonConvert.DeserializeObject<LineAuditTool.Matching.MatchConfig>(json);
                        if (loaded != null) cfg = loaded;
                    }
                }
            }
            catch (System.Exception) { }
            return cfg;
        }

        private static void EnsureMatcherInit(Autodesk.AutoCAD.DatabaseServices.Database masterDb, Editor ed)
        {
            if (_masterIdx != null && _matchCfg != null && _preview != null) return;

            try
            {
                string cacheDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "LineAuditTool");
                Directory.CreateDirectory(cacheDir);
                string cachePath = Path.Combine(cacheDir, "MasterIndex.json");

                _matchCfg = LoadConfigOrDefault();
                _preview = new LineAuditTool.Matching.PreviewRenderer();
                _masterIdx = LineAuditTool.Matching.MasterIndex.LoadOrBuild(masterDb, cachePath);

                if (ed != null) ed.WriteMessage("\n[Matcher] Init OK. Masters=" + _masterIdx.Masters.Count);
            }
            catch (System.Exception ex)
            {
                if (ed != null) ed.WriteMessage("\n[Matcher] Init failed: " + ex.Message);
                _masterIdx = null; _matchCfg = null; _preview = null; // ensure disabled if failed
            }
        }

        // ======== PUBLIC COMMAND ========
        [CommandMethod("MATCHANDREPLACECHAIRS")]
        public void MatchAndReplaceChairsFromMaster()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            RunPythonExporterOnce(ed);

            var selRes = ed.GetSelection(
                new PromptSelectionOptions { MessageForAdding = "\nSelect chair blocks to match and replace:" },
                new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "INSERT") })
            );

            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\n⛔ No blocks selected.");
                return;
            }

            // Group selected INSERTs by block name
            var blockRefsByName = new Dictionary<string, List<ObjectId>>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject sel in selRes.Value)
                {
                    if (sel == null) continue;
                    var br = tr.GetObject(sel.ObjectId, OpenMode.ForRead) as BlockReference;
                    var btr = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    string name = btr.Name;
                    if (!blockRefsByName.ContainsKey(name))
                        blockRefsByName[name] = new List<ObjectId>();
                    blockRefsByName[name].Add(sel.ObjectId);
                }
                tr.Commit();
            }

            // Open master DWG and collect block definitions
            var masterBlocks = new Dictionary<string, ObjectId>();
            using (var masterDb = new Database(false, true))
            {
                masterDb.ReadDwgFile(MasterDwgPath, FileOpenMode.OpenForReadAndAllShare, true, "");

                using (var mtr = masterDb.TransactionManager.StartTransaction())
                {
                    var mbt = (BlockTable)mtr.GetObject(masterDb.BlockTableId, OpenMode.ForRead);
                    foreach (ObjectId id in mbt)
                    {
                        var btr = (BlockTableRecord)mtr.GetObject(id, OpenMode.ForRead);
                        if (!btr.IsAnonymous && !btr.IsLayout)
                            masterBlocks[btr.Name] = id;
                    }
                    mtr.Commit();
                }

                // NEW: build/load matcher (once)
                EnsureMatcherInit(masterDb, ed);

                foreach (var kvp in blockRefsByName)
                {
                    string selectedBlock = kvp.Key;

                    // ---- Build candidate list (NEW ranker → fallback if unavailable) ----
                    List<string> suggestions = null;

                    if (_masterIdx != null && _matchCfg != null && _preview != null && _masterIdx.Masters != null && _masterIdx.Masters.Count > 0)
                    {
                        try
                        {
                            // Pick a representative source INSERT from this group
                            ObjectId sampleId = kvp.Value[0];
                            using (var trSrc = db.TransactionManager.StartTransaction())
                            {
                                var srcBr = (BlockReference)trSrc.GetObject(sampleId, OpenMode.ForRead);

                                var ranked = LineAuditTool.Matching.Scoring.RankCandidates(
                                    srcBr,
                                    trSrc,                      // read source geometry here
                                    _masterIdx.Masters,         // precomputed master entries
                                    _matchCfg,
                                    _preview                    // preview cache
                                );

                                suggestions = new List<string>(5);
                                for (int i = 0; i < ranked.Count && i < 5; i++)
                                    suggestions.Add(ranked[i].MasterName);

                                trSrc.Commit();
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("\n[Matcher] ranking failed, fallback path: " + ex.Message);
                            suggestions = null;
                        }
                    }

                    // Fallback to previous visual + name path if suggestions are null/empty
                    if (suggestions == null || suggestions.Count == 0)
                    {
                        // Preview for the SELECTED block (from ACTIVE db)
                        Bitmap selectedPreview = null;
                        try { selectedPreview = GetPreviewForActive(db, selectedBlock); } catch { }
                        bool previewUseful = IsPreviewUseful(selectedPreview);

                        if (!previewUseful)
                        {
                            ed.WriteMessage("\n(i) Selected preview looks empty → NAME-ONLY fallback.");

                            suggestions = masterBlocks.Keys
                                .Select(n => new { n, s = ComputeNameSimilarity(selectedBlock, n) })
                                .OrderByDescending(t => t.s)
                                .Take(NamePrefilterTopK)
                                .Take(5)
                                .Select(t => t.n)
                                .ToList();
                        }
                        else
                        {
                            ed.WriteMessage("\n(i) Visual-first with size gate and shape signature (fallback)...");

                            // Selected features (definition-level)
                            TryGetBlockStats(db, selectedBlock, out var selStats);
                            var selHashes = ComputeRotationalDHashes(selectedPreview); // 0/90/180/270
                            var selShape = ComputeShapeSignature(selectedPreview, SHAPE_RAYS, FG_THRESH);

                            var ranked = new List<(string name, double score)>(masterBlocks.Count);

                            foreach (var n in masterBlocks.Keys)
                            {
                                try
                                {
                                    // --- SIZE FILTER / SCORE ---
                                    TryGetBlockStats(masterDb, n, out var candStats);
                                    double sizeSim = SizeSimilarity(selStats, candStats);
                                    if (sizeSim < SIZE_GATE) continue; // keep sofas with sofas, etc.

                                    using (var candBmp = GetPreviewForMasterCandidate(masterDb, n, ed))
                                    {
                                        // --- VIS: dHash rotation-tolerant ---
                                        ulong candHash = ComputeDHash(candBmp);
                                        int minHd = 64;
                                        for (int i = 0; i < selHashes.Length; i++)
                                        {
                                            int hd = HammingDistance(selHashes[i], candHash);
                                            if (hd < minHd) minHd = hd;
                                        }
                                        double vis = 1.0 - (double)minHd / 64.0;

                                        // --- SHAPE: radial signature correlation ---
                                        double shapeSim = 0.0;
                                        try
                                        {
                                            var candShape = ComputeShapeSignature(candBmp, SHAPE_RAYS, FG_THRESH);
                                            shapeSim = BestCircularCorrelation(selShape, candShape);
                                        }
                                        catch { }

                                        double nameSim = NormalizeNameScore(ComputeNameSimilarity(selectedBlock, n));
                                        double total = W_VIS * vis + W_SHAPE * shapeSim + W_SIZE * sizeSim + W_NAME * nameSim;
                                        ranked.Add((n, total));
                                    }
                                }
                                catch { }
                            }

                            suggestions = ranked
                                .OrderByDescending(r => r.score)
                                .Take(5)
                                .Select(r => r.name)
                                .ToList();
                        }

                        selectedPreview?.Dispose();
                    }

                    // ---- UI picker ----
                    string chosenMasterName = ShowPreviewSelectionForm(selectedBlock, suggestions, masterDb, ed);
                    if (string.IsNullOrEmpty(chosenMasterName))
                    {
                        ed.WriteMessage($"\n⏭️ Skipping '{selectedBlock}'...");
                        continue;
                    }

                    ed.WriteMessage($"\n✅ Replacing '{selectedBlock}' with '{chosenMasterName}'...");

                    // Layer/color sync from an instance of the chosen master block
                    string masterRefLayer = "0";
                    Autodesk.AutoCAD.Colors.Color masterLayerColor =
                        Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByAci, 7);

                    using (var mtr2 = masterDb.TransactionManager.StartTransaction())
                    {
                        var mbt = (BlockTable)mtr2.GetObject(masterDb.BlockTableId, OpenMode.ForRead);
                        var ms = (BlockTableRecord)mtr2.GetObject(mbt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                        foreach (ObjectId entId in ms)
                        {
                            var br = mtr2.GetObject(entId, OpenMode.ForRead) as BlockReference;
                            if (br == null) continue;
                            var def = (BlockTableRecord)mtr2.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                            if (def.Name != chosenMasterName) continue;

                            masterRefLayer = br.Layer;
                            var lt = (LayerTable)mtr2.GetObject(masterDb.LayerTableId, OpenMode.ForRead);
                            if (lt.Has(masterRefLayer))
                            {
                                var ltr = (LayerTableRecord)mtr2.GetObject(lt[masterRefLayer], OpenMode.ForRead);
                                masterLayerColor = ltr.Color;
                            }
                            break;
                        }
                        mtr2.Commit();
                    }

                    // Ensure layer exists & color matches in active drawing
                    using (var tr2 = db.TransactionManager.StartTransaction())
                    {
                        var lt = (LayerTable)tr2.GetObject(db.LayerTableId, OpenMode.ForRead);
                        if (!lt.Has(masterRefLayer))
                        {
                            lt.UpgradeOpen();
                            var newLayer = new LayerTableRecord { Name = masterRefLayer, Color = masterLayerColor };
                            lt.Add(newLayer);
                            tr2.AddNewlyCreatedDBObject(newLayer, true);
                        }
                        else
                        {
                            var existing = (LayerTableRecord)tr2.GetObject(lt[masterRefLayer], OpenMode.ForWrite);
                            existing.Color = masterLayerColor;
                        }
                        tr2.Commit();
                    }

                    // Bring chosen definition into current drawing (replace if exists)
                    var idMap = new IdMapping();
                    masterDb.WblockCloneObjects(
                        new ObjectIdCollection(new[] { masterBlocks[chosenMasterName] }),
                        db.BlockTableId,
                        idMap,
                        DuplicateRecordCloning.Replace,
                        false
                    );

                    // Replace selected instances
                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        // Preload candidate preview once for orientation estimation (NCC)
                        Bitmap candPreview = GetPreviewForMasterCandidate(masterDb, chosenMasterName, ed);

                        foreach (ObjectId id in kvp.Value)
                        {
                            var oldBr = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);

                            // Capture transform
                            Point3d oldPos = oldBr.Position;
                            double oldRot = oldBr.Rotation;
                            Scale3d oldScale = oldBr.ScaleFactors;

                            // Definitions
                            var oldDef = (BlockTableRecord)tr.GetObject(oldBr.BlockTableRecord, OpenMode.ForRead);
                            var newDef = (BlockTableRecord)tr.GetObject(bt[chosenMasterName], OpenMode.ForRead);

                            // Pivots
                            Point3d oldLocalPivot = GetBlockPivot(oldDef, tr);
                            Point3d newLocalPivot = GetBlockPivot(newDef, tr);

                            // ---- ORIENTATION FIX (multiscale NCC) ----
                            double deltaAngle = 0.0; // radians
                            try
                            {
                                Bitmap selectedDefPreviewForAngle = GenerateBlockPreview(db, selectedBlock);
                                if (selectedDefPreviewForAngle != null && candPreview != null)
                                {
                                    double bestDeg = EstimateDeltaAngleByImageNCC(
                                        selectedDefPreviewForAngle,   // selected DEF preview
                                        candPreview                   // candidate DEF preview
                                    );
                                    // rotate candidate by -bestDeg to align to selected definition
                                    deltaAngle = -bestDeg * Math.PI / 180.0;
                                }
                            }
                            catch { /* fall through silently */ }

                            // Old pivot in WORLD
                            Matrix3d oldXform = oldBr.BlockTransform;
                            Point3d oldPivotWorld = oldLocalPivot.TransformBy(oldXform);

                            // Remove old
                            oldBr.Erase();

                            // Insert new with old rotation + orientation correction
                            var newBr = new BlockReference(oldPos, bt[chosenMasterName])
                            {
                                Rotation = oldRot + deltaAngle,
                                ScaleFactors = oldScale,
                                Layer = masterRefLayer
                            };
                            ms.AppendEntity(newBr);
                            tr.AddNewlyCreatedDBObject(newBr, true);

                            // Align pivots
                            Matrix3d newXform = newBr.BlockTransform;
                            Point3d newPivotWorld = newLocalPivot.TransformBy(newXform);
                            Vector3d delta = oldPivotWorld - newPivotWorld;
                            if (!delta.IsZeroLength())
                                newBr.TransformBy(Matrix3d.Displacement(delta));
                        }

                        candPreview?.Dispose();
                        tr.Commit();
                    }
                }
            }
        }

        // ======== NAME similarity (fallback path) ========
        private int ComputeNameSimilarity(string a, string b)
        {
            a = a.ToLower(); b = b.ToLower();
            int score = 0;
            if (a == b) score += 100;
            if (a.Contains(b) || b.Contains(a)) score += 50;
            score += a.Intersect(b).Count();
            return score;
        }

        private static double NormalizeNameScore(int s)
        {
            double v = s / 200.0;
            if (v < 0) v = 0;
            if (v > 1) v = 1;
            return v;
        }

        // ======== PREVIEW (vector fallback; used by picker + NCC) ========
        private Bitmap GenerateBlockPreview(Database db, string blockName)
        {
            var bmp = new Bitmap(128, 128);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(DrawingColor.White);
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    if (!bt.Has(blockName)) return PreparePreviewBitmap(bmp, 128, 8);
                    var btr = (BlockTableRecord)tr.GetObject(bt[blockName], OpenMode.ForRead);

                    Extents3d? extOpt = null;
                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        try
                        {
                            var e = ent.GeometricExtents;
                            extOpt = extOpt == null ? (Extents3d?)e : UnionExtents(extOpt.Value, e);
                        }
                        catch { }
                    }
                    if (extOpt == null) return PreparePreviewBitmap(bmp, 128, 8);
                    var ext = extOpt.Value;

                    double width = ext.MaxPoint.X - ext.MinPoint.X;
                    double height = ext.MaxPoint.Y - ext.MinPoint.Y;
                    if (width <= 1e-9 || height <= 1e-9) return PreparePreviewBitmap(bmp, 128, 8);

                    double target = 116.0;
                    double scale = Math.Min(target / width, target / height);
                    double offX = (bmp.Width - width * scale) / 2.0;
                    double offY = (bmp.Height - height * scale) / 2.0;

                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        DrawEntitySmart(g, ent, ext, scale, offX, offY, bmp.Height);
                    }
                    tr.Commit();
                }
            }
            return PreparePreviewBitmap(bmp, 128, 8);
        }

        private void DrawEntitySmart(Graphics g, Entity ent, Extents3d ext,
                                     double scale, double offsetX, double offsetY, int canvasH)
        {
            try
            {
                if (ent is Polyline pl)
                {
                    for (int i = 0; i < pl.NumberOfVertices - 1; i++)
                    {
                        var p1 = pl.GetPoint2dAt(i);
                        var p2 = pl.GetPoint2dAt(i + 1);
                        double bulge = pl.GetBulgeAt(i);
                        if (bulge != 0) DrawArcSegment(g, p1, p2, bulge, ext, scale, offsetX, offsetY, canvasH);
                        else DrawLineSegment(g, p1.X, p1.Y, p2.X, p2.Y, ext, scale, offsetX, offsetY, canvasH);
                    }
                    return;
                }
                if (ent is Line ln)
                {
                    DrawLineSegment(g, ln.StartPoint.X, ln.StartPoint.Y, ln.EndPoint.X, ln.EndPoint.Y, ext, scale, offsetX, offsetY, canvasH);
                    return;
                }
                if (ent is Arc arc)
                {
                    var r = ToScreenRectangle(arc.Center, arc.Radius, ext, scale, offsetX, offsetY, canvasH);
                    float a0 = (float)RadiansToDegrees(arc.StartAngle);
                    float a1 = (float)RadiansToDegrees(arc.EndAngle - arc.StartAngle);
                    g.DrawArc(Pens.Black, r, a0, a1);
                    return;
                }
                if (ent is Circle circ)
                {
                    var r = ToScreenRectangle(circ.Center, circ.Radius, ext, scale, offsetX, offsetY, canvasH);
                    g.DrawEllipse(Pens.Black, r);
                    return;
                }
                if (ent is Ellipse el)
                {
                    var r = ToScreenRectangle(el.Center, el.MajorRadius, ext, scale, offsetX, offsetY, canvasH);
                    g.DrawEllipse(Pens.Black, r);
                    return;
                }
                if (ent is Spline sp)
                {
                    int n = 30;
                    Point3d prev = sp.GetPointAtParameter(sp.StartParam);
                    for (int i = 1; i <= n; i++)
                    {
                        double t = sp.StartParam + (sp.EndParam - sp.StartParam) * i / n;
                        Point3d cur = sp.GetPointAtParameter(t);
                        DrawLineSegment(g, prev.X, prev.Y, cur.X, cur.Y, ext, scale, offsetX, offsetY, canvasH);
                        prev = cur;
                    }
                    return;
                }
                if (ent is BlockReference br)
                {
                    var objs = new DBObjectCollection();
                    br.Explode(objs);
                    foreach (DBObject dob in objs)
                    {
                        var e = dob as Entity;
                        if (e != null) DrawEntitySmart(g, e, ext, scale, offsetX, offsetY, canvasH);
                        dob.Dispose();
                    }
                    return;
                }
            }
            catch { }
        }

        private void DrawLineSegment(Graphics g, double x1, double y1, double x2, double y2,
                                     Extents3d ext, double scale, double offsetX, double offsetY, int canvasH)
        {
            g.DrawLine(Pens.Black,
                (float)((x1 - ext.MinPoint.X) * scale + offsetX),
                (float)(canvasH - ((y1 - ext.MinPoint.Y) * scale + offsetY)),
                (float)((x2 - ext.MinPoint.X) * scale + offsetX),
                (float)(canvasH - ((y2 - ext.MinPoint.Y) * scale + offsetY)));
        }

        private void DrawArcSegment(Graphics g, Point2d p1, Point2d p2, double bulge,
                                    Extents3d ext, double scale, double offsetX, double offsetY, int canvasH)
        {
            double b = Math.Abs(bulge);
            double ang = 4 * Math.Atan(b);
            double chord = p1.GetDistanceTo(p2);
            double radius = chord / (2 * Math.Sin(ang / 2));

            var dir = (p2 - p1).GetNormal();
            var normal = new Vector2d(-dir.Y, dir.X);
            double h = radius * Math.Cos(ang / 2);
            var mid = new Point2d((p1.X + p2.X) / 2, (p1.Y + p2.Y) / 2);
            var center = mid + normal * (b > 0 ? h : -h);

            double start = Math.Atan2(p1.Y - center.Y, p1.X - center.X);
            double end = Math.Atan2(p2.Y - center.Y, p2.X - center.X);
            double sweep = RadiansToDegrees(end - start);
            if (sweep < 0) sweep += 360;

            var rect = ToScreenRectangle(new Point3d(center.X, center.Y, 0), radius, ext, scale, offsetX, offsetY, canvasH);
            using (var path = new System.Drawing.Drawing2D.GraphicsPath())
            {
                path.AddArc(rect, (float)RadiansToDegrees(start), (float)sweep);
                g.DrawPath(Pens.Black, path);
            }
        }

        private double RadiansToDegrees(double r) => r * (180.0 / Math.PI);

        private RectangleF ToScreenRectangle(Point3d center, double radius, Extents3d ext,
                                             double scale, double offsetX, double offsetY, int canvasH)
        {
            float cx = (float)((center.X - ext.MinPoint.X) * scale + offsetX);
            float cy = (float)(canvasH - ((center.Y - ext.MinPoint.Y) * scale + offsetY));
            float rr = (float)(radius * scale);
            return new RectangleF(cx - rr, cy - rr, rr * 2, rr * 2);
        }

        private Extents3d UnionExtents(Extents3d a, Extents3d b)
        {
            var min = new Point3d(Math.Min(a.MinPoint.X, b.MinPoint.X), Math.Min(a.MinPoint.Y, b.MinPoint.Y), 0);
            var max = new Point3d(Math.Max(a.MaxPoint.X, b.MaxPoint.X), Math.Max(a.MaxPoint.Y, b.MaxPoint.Y), 0);
            return new Extents3d(min, max);
        }

        // ======== PIVOT ========
        private static Point3d GetBlockPivot(BlockTableRecord btr, Transaction tr)
        {
            Extents3d? extOpt = null;
            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                try
                {
                    var e = ent.GeometricExtents;
                    extOpt = extOpt == null
                        ? (Extents3d?)e
                        : new Extents3d(
                            new Point3d(Math.Min(extOpt.Value.MinPoint.X, e.MinPoint.X),
                                        Math.Min(extOpt.Value.MinPoint.Y, e.MinPoint.Y), 0),
                            new Point3d(Math.Max(extOpt.Value.MaxPoint.X, e.MaxPoint.X),
                                        Math.Max(extOpt.Value.MaxPoint.Y, e.MaxPoint.Y), 0));
                }
                catch { }
            }

            if (extOpt.HasValue)
            {
                var ex = extOpt.Value;
                return new Point3d(
                    (ex.MinPoint.X + ex.MaxPoint.X) * 0.5,
                    (ex.MinPoint.Y + ex.MaxPoint.Y) * 0.5,
                    0.0);
            }
            return Point3d.Origin;
        }

        // ======== SIZE & SHAPE FEATURES (fallback path) ========
        private struct BlockStats
        {
            public double Width, Height, Area, Aspect;
            public bool IsValid => Width > 1e-9 && Height > 1e-9;
        }

        private bool TryGetBlockStats(Database db, string blockName, out BlockStats stats)
        {
            stats = new BlockStats();
            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    if (!bt.Has(blockName)) return false;
                    var btr = (BlockTableRecord)tr.GetObject(bt[blockName], OpenMode.ForRead);

                    Extents3d? extOpt = null;
                    foreach (ObjectId id in btr)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        try
                        {
                            var e = ent.GeometricExtents;
                            extOpt = extOpt == null ? (Extents3d?)e : UnionExtents(extOpt.Value, e);
                        }
                        catch { }
                    }
                    if (extOpt == null) return false;
                    var ext = extOpt.Value;

                    double w = ext.MaxPoint.X - ext.MinPoint.X;
                    double h = ext.MaxPoint.Y - ext.MinPoint.Y;
                    stats.Width = w;
                    stats.Height = h;
                    stats.Area = Math.Max(1e-9, w * h);
                    stats.Aspect = w / Math.Max(1e-9, h);
                    return stats.IsValid;
                }
            }
            catch { return false; }
        }

        private static double SizeSimilarity(BlockStats a, BlockStats b)
        {
            if (!a.IsValid || !b.IsValid) return 0.0;

            double aspectSim = Math.Exp(-Math.Abs(Math.Log((a.Aspect + 1e-9) / (b.Aspect + 1e-9))));
            double areaSim = Math.Exp(-Math.Abs(Math.Log((a.Area + 1e-9) / (b.Area + 1e-9))));
            if (aspectSim < 0) aspectSim = 0; if (aspectSim > 1) aspectSim = 1;
            if (areaSim < 0) areaSim = 0; if (areaSim > 1) areaSim = 1;

            return 0.6 * aspectSim + 0.4 * areaSim; // 0..1
        }

        private static double[] ComputeShapeSignature(Bitmap bmp, int rays, byte fg)
        {
            int W = bmp.Width, H = bmp.Height;
            double cx = (W - 1) * 0.5, cy = (H - 1) * 0.5;
            double rMax = Math.Min(cx, cy);

            var sig = new double[rays];
            for (int i = 0; i < rays; i++)
            {
                double ang = (2.0 * Math.PI * i) / rays;
                double dx = Math.Cos(ang), dy = Math.Sin(ang);

                double rHit = 0.0;
                for (double r = 0; r <= rMax; r += 1.0)
                {
                    int x = (int)Math.Round(cx + dx * r);
                    int y = (int)Math.Round(cy - dy * r);
                    if (x < 0 || x >= W || y < 0 || y >= H) break;
                    var c = bmp.GetPixel(x, y);
                    byte lum = (byte)(0.299 * c.R + 0.587 * c.G + 0.114 * c.B);
                    if (lum < fg) { rHit = r; break; }
                }
                sig[i] = rHit;
            }

            double max = sig.Max();
            if (max < 1e-9) return sig;
            for (int i = 0; i < rays; i++) sig[i] /= max;
            return sig;
        }

        private static double BestCircularCorrelation(double[] a, double[] b)
        {
            if (a.Length != b.Length || a.Length == 0) return 0.0;
            int n = a.Length;
            double best = -1.0;

            for (int shift = 0; shift < n; shift++)
            {
                double num = 0, denA = 0, denB = 0;
                for (int i = 0; i < n; i++)
                {
                    double ai = a[i];
                    double bi = b[(i + shift) % n];
                    num += ai * bi;
                    denA += ai * ai;
                    denB += bi * bi;
                }
                double corr = (denA > 0 && denB > 0) ? (num / Math.Sqrt(denA * denB)) : 0.0;
                if (corr > best) best = corr;
            }
            if (best < 0) best = 0;
            if (best > 1) best = 1;
            return best;
        }

        // ======== Thumbnail & Visual Hash helpers (fallback path) ========
        private static Bitmap PreparePreviewBitmap(Bitmap src, int box = 128, int margin = 8)
        {
            var dst = new Bitmap(box, box);
            using (var g = Graphics.FromImage(dst))
            {
                g.Clear(DrawingColor.White);
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                double maxW = box - 2 * margin, maxH = box - 2 * margin;
                double scale = Math.Min(maxW / src.Width, maxH / src.Height);
                int w = Math.Max(1, (int)Math.Round(src.Width * scale));
                int h = Math.Max(1, (int)Math.Round(src.Height * scale));
                int x = (box - w) / 2, y = (box - h) / 2;

                g.DrawImage(src, new Rectangle(x, y, w, h));
            }
            return dst;
        }

        private static Bitmap LoadBitmapClone(string path)
        {
            using (var tmp = new Bitmap(path))
            using (var clone = new Bitmap(tmp))
                return PreparePreviewBitmap(clone, 128, 8);
        }

        private Bitmap GetPreviewForMasterCandidate(Database masterDb, string blockName, Editor ed = null)
        {
            string path = FindPythonPreviewPath(blockName, ed);
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
                return LoadBitmapClone(path);
            return GenerateBlockPreview(masterDb, blockName);
        }

        private Bitmap GetPreviewForActive(Database activeDb, string blockName) =>
            GenerateBlockPreview(activeDb, blockName);

        private bool IsPreviewUseful(Bitmap bmp)
        {
            if (bmp == null) return false;
            int w = bmp.Width, h = bmp.Height;
            int dark = 0, total = (w * h) / 4;
            for (int y = 0; y < h; y += 2)
                for (int x = 0; x < w; x += 2)
                {
                    var c = bmp.GetPixel(x, y);
                    if (c.R < 240 || c.G < 240 || c.B < 240) dark++;
                }
            double ratio = total > 0 ? (double)dark / total : 0.0;
            return ratio > 0.01;
        }

        // ----- dHash -----
        private static ulong ComputeDHash(Bitmap bmp)
        {
            const int W = 9, H = 8;
            using (var small = new Bitmap(W, H))
            using (var g = Graphics.FromImage(small))
            {
                g.DrawImage(bmp, new Rectangle(0, 0, W, H));
                ulong hash = 0UL;
                int bit = 0;
                for (int y = 0; y < H; y++)
                    for (int x = 0; x < W - 1; x++)
                    {
                        var c1 = small.GetPixel(x, y);
                        var c2 = small.GetPixel(x + 1, y);
                        int l1 = (int)(0.299 * c1.R + 0.587 * c1.G + 0.114 * c1.B);
                        int l2 = (int)(0.299 * c2.R + 0.587 * c2.G + 0.114 * c2.B);
                        if (l1 > l2) hash |= (1UL << bit);
                        bit++;
                    }
                return hash;
            }
        }

        private static ulong[] ComputeRotationalDHashes(Bitmap bmp)
        {
            var hashes = new ulong[4];
            Bitmap tmp = (Bitmap)bmp.Clone();
            for (int i = 0; i < 4; i++)
            {
                hashes[i] = ComputeDHash(tmp);
                tmp.RotateFlip(RotateFlipType.Rotate90FlipNone);
            }
            tmp.Dispose();
            return hashes;
        }

        private static int HammingDistance(ulong a, ulong b)
        {
            ulong x = a ^ b;
            int c = 0;
            while (x != 0) { x &= (x - 1); c++; }
            return c;
        }

        // ======== Precise orientation estimator (multiscale NCC) ========
        // Returns θ degrees so that rotate(selectedDef, θ) best matches candidateDef.
        private static double EstimateDeltaAngleByImageNCC(Bitmap selectedDefPreview, Bitmap candidateDefPreview)
        {
            // Precompute normalized candidate vector (downsampled)
            var candVec = ToNormalizedGrayVector(candidateDefPreview, 2); // 64x64 = 4096

            double bestAngle = 0;
            double bestScore = double.NegativeInfinity;

            Action<double, double, double> scan = (center, span, step) =>
            {
                double start = center - span;
                double end = center + span + 1e-9;
                for (double a = start; a < end; a += step)
                {
                    using (var rot = RotateIntoBox(selectedDefPreview, a))
                    {
                        var v = ToNormalizedGrayVector(rot, 2);
                        double score = Dot(v, candVec); // NCC since both are mean-0, std-1
                        if (score > bestScore)
                        {
                            bestScore = score;
                            bestAngle = NormalizeDeg(a);
                        }
                    }
                }
            };

            // Pass 1: coarse
            scan(180, 180, 10.0);
            // Pass 2: fine around best
            scan(bestAngle, 12, 1.0);
            // Pass 3: micro around best
            scan(bestAngle, 2, 0.25);

            return bestAngle; // degrees
        }

        private static double NormalizeDeg(double a)
        {
            while (a >= 360.0) a -= 360.0;
            while (a < 0.0) a += 360.0;
            return a;
        }

        private static Bitmap RotateIntoBox(Bitmap src, double angleDeg)
        {
            int W = 128, H = 128;
            var dst = new Bitmap(W, H);
            using (var g = Graphics.FromImage(dst))
            {
                g.Clear(DrawingColor.White);
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                g.TranslateTransform(W / 2f, H / 2f);
                g.RotateTransform((float)angleDeg);
                g.TranslateTransform(-W / 2f, -H / 2f);
                g.DrawImage(src, new Rectangle(0, 0, W, H));
            }
            return dst;
        }

        private static double[] ToNormalizedGrayVector(Bitmap bmp, int stride)
        {
            // Downsample by 'stride' (2 => 64x64). Convert to grayscale, normalize mean/std.
            int W = bmp.Width, H = bmp.Height;
            int w = W / stride, h = H / stride;
            var v = new double[w * h];

            int k = 0; double sum = 0, sum2 = 0;
            for (int y = 0; y < H; y += stride)
            {
                for (int x = 0; x < W; x += stride)
                {
                    var c = bmp.GetPixel(x, y);
                    double g = 0.299 * c.R + 0.587 * c.G + 0.114 * c.B;
                    v[k++] = g;
                    sum += g;
                    sum2 += g * g;
                }
            }
            int n = v.Length;
            double mean = sum / n;
            double var = Math.Max(1e-9, (sum2 / n) - mean * mean);
            double inv = 1.0 / Math.Sqrt(var);
            for (int i = 0; i < n; i++) v[i] = (v[i] - mean) * inv;
            return v;
        }

        private static double Dot(double[] a, double[] b)
        {
            double s = 0; int n = a.Length;
            for (int i = 0; i < n; i++) s += a[i] * b[i];
            return s / n; // normalize by length so scores are comparable
        }

        // ======== Python preview integration (for picker thumbnails) ========
        private static string DrawingFolderFromMaster
        {
            get { try { return Path.GetFileNameWithoutExtension(MasterDwgPath); } catch { return "M1"; } }
        }

        private static string InstancesDirFullPath(string root) =>
            Path.Combine(root, DrawingFolderFromMaster, InstancesDirName);

        private static string AutoClustersDirFullPath(string root) =>
            Path.Combine(root, DrawingFolderFromMaster, AutoClustersDirName);

        private static string MakeSafe(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Unnamed";
            foreach (var c in "<>:\"/\\|?*".ToCharArray()) name = name.Replace(c, '_');
            return name.Trim();
        }

        private static bool ContainsCI(string haystack, string needle) =>
            haystack != null && needle != null &&
            haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;

        private static string NormalizeToken(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            foreach (var c in "<>:\"/\\|?*".ToCharArray()) s = s.Replace(c, '_');
            return s.Replace(' ', '_').ToLowerInvariant();
        }

        private static IEnumerable<string> GetFilesMultiPattern(string dir, params string[] patterns)
        {
            if (!Directory.Exists(dir)) yield break;
            foreach (var pat in patterns)
            {
                string[] arr = new string[0];
                try { arr = Directory.GetFiles(dir, pat, SearchOption.TopDirectoryOnly); }
                catch { }
                for (int j = 0; j < arr.Length; j++) yield return arr[j];
            }
        }

        private static string FindPythonPreviewPath(string blockName, Editor ed = null)
        {
            try
            {
                var safe = MakeSafe(blockName);
                var key = NormalizeToken(safe);

                string instDir = InstancesDirFullPath(PythonExportsRoot);
                string cluDir = AutoClustersDirFullPath(PythonExportsRoot);

                ed?.WriteMessage($"\n[Preview] Key='{key}'  INST='{instDir}'  CLU='{cluDir}'");

                var instExact = GetFilesMultiPattern(instDir, safe + "__*.jpg").FirstOrDefault();
                if (!string.IsNullOrEmpty(instExact)) return instExact;

                var instLoose = GetFilesMultiPattern(instDir, "*.jpg")
                    .FirstOrDefault(p => ContainsCI(Path.GetFileNameWithoutExtension(p), key));
                if (!string.IsNullOrEmpty(instLoose)) return instLoose;

                var cluExact = GetFilesMultiPattern(cluDir, safe + "__cluster*.jpg").FirstOrDefault();
                if (!string.IsNullOrEmpty(cluExact)) return cluExact;

                var cluLoose1 = GetFilesMultiPattern(cluDir, "*.jpg")
                    .FirstOrDefault(p => ContainsCI(Path.GetFileNameWithoutExtension(p), key));
                if (!string.IsNullOrEmpty(cluLoose1)) return cluLoose1;

                var keyNoUnderscore = key.Replace("_", "");
                var cluLoose2 = GetFilesMultiPattern(cluDir, "*.jpg")
                    .FirstOrDefault(p =>
                    {
                        var name = NormalizeToken(Path.GetFileNameWithoutExtension(p)).Replace("_", "");
                        return name.Contains(keyNoUnderscore);
                    });
                if (!string.IsNullOrEmpty(cluLoose2)) return cluLoose2;
            }
            catch { }
            return null;
        }

        // ======== Picker UI ========
        private string ShowPreviewSelectionForm(string selectedBlock, List<string> suggestions, Database masterDb, Editor ed)
        {
            RunPythonExporterOnce(ed);

            var form = new Form { Text = "Suggestions for " + selectedBlock, Width = 560, Height = 440 };
            var listView = new ListView { Dock = DockStyle.Fill, View = View.LargeIcon, MultiSelect = false };
            var imageList = new ImageList { ImageSize = new Size(128, 128), ColorDepth = ColorDepth.Depth32Bit };
            listView.LargeImageList = imageList;

            using (var tr = masterDb.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(masterDb.BlockTableId, OpenMode.ForRead);

                foreach (var name in suggestions)
                {
                    if (!bt.Has(name)) continue;

                    ed?.WriteMessage("\n[Preview] Trying: " + name);

                    string chosenPath = FindPythonPreviewPath(name, ed);
                    Bitmap bmp;

                    if (!string.IsNullOrEmpty(chosenPath) && File.Exists(chosenPath))
                    {
                        ed?.WriteMessage("  -> from file: " + chosenPath);
                        bmp = LoadBitmapClone(chosenPath);
                    }
                    else
                    {
                        ed?.WriteMessage("  -> no file, using vector preview.");
                        bmp = GenerateBlockPreview(masterDb, name);
                    }

                    imageList.Images.Add(name, bmp);
                    listView.Items.Add(new ListViewItem(name, name));
                }
                tr.Commit();
            }

            form.Controls.Add(listView);

            string selected = null;
            var ok = new Button { Text = "OK", Dock = DockStyle.Bottom };
            ok.Click += (s, e) =>
            {
                if (listView.SelectedItems.Count > 0)
                    selected = listView.SelectedItems[0].Text;
                form.Close();
            };
            form.Controls.Add(ok);

            form.ShowDialog();
            return selected;
        }

        // ======== EXPOSED ========
        public static void RunPythonExporterOnce(Editor ed = null)
        {
            if (_exporterRunThisSession) return;
            _exporterRunThisSession = true;

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = PythonExe,
                    Arguments = $"\"{PythonExporterScript}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(PythonExporterScript)
                };

                var p = Process.Start(psi);
                if (p != null)
                {
                    if (!p.WaitForExit(PythonTimeoutMs)) { try { p.Kill(); } catch { } }
                    p.Dispose();
                }

                ed?.WriteMessage("\n(i) Python preview export attempted (run-once).");
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage("\n(!) Python export failed: " + ex.Message);
            }
        }
    }
}
