// File: LineAuditCommand.cs
// Commands:
//   POLY2TRI_VALIDATE       → find & mark CDT-breaking issues
//   POLY2TRI_CLEAR_ERRORS   → delete all markers/labels on _TRI_ERRORS
//   CURVES2LWP / C2P        → curve→LWP (smart-join, dup + near-twin filters)
//   C2PD                    → densify-only (no smoothing)
//
// Notes:
// - Works with LWPOLYLINE/2D/3D (open/closed), CIRCLE, ARC, SPLINE (tessellated)
// - Markers: red circles + DBText on layer _TRI_ERRORS (reduced size)

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

// Register BOTH command classes in this one assembly:
[assembly: CommandClass(typeof(LineAuditTool.BlockMatcher))]
[assembly: CommandClass(typeof(LineAuditTool.CurveToPolyline73))]

namespace LineAuditTool
{
    // =================================================================
    // =============== POLY2TRI VALIDATOR + CLEAR MARKS =================
    // =================================================================
    public class BlockMatcher
    {
        // ======== CONFIG ========
        private const string ErrorLayer = "_TRI_ERRORS";

        // Geometry tolerances
        private const double Eps = 1e-6;
        private const double NearDupEps = 1e-4;
        private const double CollinearDeg = 0.5;

        // Marker/label sizing (reduced)
        private const double MarkerRadius = 6.0;   // was 20
        private const double TextHeight = 2.2;     // fixed height
        private const double LabelOffset = 8.0;    // XY offset from marker center

        // Tessellation
        private const int CircleSides = 64;
        private const int ArcSides = 32;
        private const int SplineSamples = 64;

        // Optional CDT probe (compile-time guarded)
        private const bool USE_POLY2TRI = false;
        // =========================

        private struct Loop
        {
            public List<Point2d> Pts;
            public bool IsClosed;
            public string Kind;
        }

        private struct Issue
        {
            public Point2d Position;
            public string Message;
            public Issue(Point2d p, string m) { Position = p; Message = m; }
        }
        private class IssueComparer : IEqualityComparer<Issue>
        {
            public bool Equals(Issue a, Issue b) =>
                a.Message == b.Message && a.Position.GetDistanceTo(b.Position) < 1e-6;
            public int GetHashCode(Issue x) => x.Message.GetHashCode();
        }

        // -------------------- MAIN: VALIDATE --------------------
        [CommandMethod("POLY2TRI_VALIDATE")]
        public void ValidatePolylines()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                var pko = new PromptKeywordOptions("\nValidate shapes: ");
                pko.Keywords.Add("Selection");
                pko.Keywords.Add("Layer");
                pko.Keywords.Default = "Selection";
                var pkr = ed.GetKeywords(pko);
                if (pkr.Status != PromptStatus.OK) return;

                var ids = new List<ObjectId>();
                if (pkr.StringResult == "Selection")
                {
                    var psel = new PromptSelectionOptions();
                    psel.MessageForAdding = "\nSelect shapes (LW/2D/3D polyline, circle, arc, spline): ";
                    var filt = new SelectionFilter(new[]
                    {
                        new TypedValue((int)DxfCode.Start, "LWPOLYLINE,POLYLINE,CIRCLE,ARC,SPLINE")
                    });
                    var res = ed.GetSelection(psel, filt);
                    if (res.Status != PromptStatus.OK) return;
                    ids.AddRange(res.Value.GetObjectIds());
                }
                else
                {
                    var pr = ed.GetString("\nEnter layer name to scan: ");
                    if (pr.Status != PromptStatus.OK) return;
                    string layerName = pr.StringResult;

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                        foreach (ObjectId id in ms)
                        {
                            var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if (ent == null) continue;
                            if (!ent.Layer.Equals(layerName, StringComparison.OrdinalIgnoreCase)) continue;

                            if (ent is Polyline || ent is Polyline2d || ent is Polyline3d ||
                                ent is Circle || ent is Arc || ent is Spline)
                                ids.Add(id);
                        }
                        tr.Commit();
                    }
                }

                if (ids.Count == 0)
                {
                    ed.WriteMessage("\nNo supported entities found.");
                    return;
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    EnsureLayer(db, tr, ErrorLayer, Color.FromRgb(255, 64, 64));
                    tr.Commit();
                }

                int totalIssues = 0, totalLoops = 0;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    foreach (var id in ids)
                    {
                        var ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
                        var loops = ExtractLoops(ent);

                        foreach (var loop in loops)
                        {
                            totalLoops++;
                            var issues = ValidateLoop(loop);

                            if (USE_POLY2TRI && loop.IsClosed)
                                issues.AddRange(TryPoly2Tri(loop.Pts));

                            if (issues.Count > 0)
                            {
                                totalIssues += issues.Count;
                                DropMarkers(tr, db, issues);
                                Report(ed, ent.Handle.ToString(), loop.Kind, issues);
                            }
                        }
                    }
                    tr.Commit();
                }

                ed.WriteMessage($"\nChecked {totalLoops} loop(s). Marked {totalIssues} issue(s) on layer {ErrorLayer}.");
                ed.WriteMessage("\nTip: isolate layer _TRI_ERRORS, fix points, then run POLY2TRI_CLEAR_ERRORS to clean markers.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[POLY2TRI_VALIDATE] Error: {ex.Message}");
            }
        }

        // -------------------- CLEAN: CLEAR ERRORS --------------------
        [CommandMethod("POLY2TRI_CLEAR_ERRORS")]
        public void ClearErrorMarkersCmd()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            try
            {
                int deleted = 0;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (!lt.Has(ErrorLayer))
                    {
                        ed.WriteMessage($"\nNothing to clear: layer {ErrorLayer} does not exist.");
                        tr.Commit();
                        return;
                    }

                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    var toErase = new List<Entity>();
                    foreach (ObjectId id in ms)
                    {
                        var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        if (ent.Layer.Equals(ErrorLayer, StringComparison.OrdinalIgnoreCase))
                        {
                            ent.UpgradeOpen();
                            toErase.Add(ent);
                        }
                    }

                    foreach (var ent in toErase)
                    {
                        ent.Erase();
                        deleted++;
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\nPOLY2TRI_CLEAR_ERRORS: deleted {deleted} object(s) on {ErrorLayer}.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[POLY2TRI_CLEAR_ERRORS] Error: {ex.Message}");
            }
        }

        // ---------- Extract tessellated loops from supported entities ----------
        private List<Loop> ExtractLoops(Entity ent)
        {
            var result = new List<Loop>();

            if (ent is Polyline pl)
            {
                var pts = new List<Point2d>();
                for (int i = 0; i < pl.NumberOfVertices; i++)
                    pts.Add(pl.GetPoint2dAt(i));
                if (pts.Count > 1 && pts.First().GetDistanceTo(pts.Last()) < NearDupEps)
                    pts.RemoveAt(pts.Count - 1);
                result.Add(new Loop { Pts = pts, IsClosed = pl.Closed, Kind = "LWPolyline" });
            }
            else if (ent is Polyline2d p2)
            {
                var pts = new List<Point2d>();
                foreach (ObjectId vId in p2)
                {
                    using (var v = (Vertex2d)vId.GetObject(OpenMode.ForRead))
                        pts.Add(new Point2d(v.Position.X, v.Position.Y));
                }
                if (pts.Count > 1 && pts.First().GetDistanceTo(pts.Last()) < NearDupEps)
                    pts.RemoveAt(pts.Count - 1);
                result.Add(new Loop { Pts = pts, IsClosed = p2.Closed, Kind = "Polyline2d" });
            }
            else if (ent is Polyline3d p3)
            {
                var pts = new List<Point2d>();
                foreach (ObjectId vId in p3)
                {
                    using (var v = (PolylineVertex3d)vId.GetObject(OpenMode.ForRead))
                        pts.Add(new Point2d(v.Position.X, v.Position.Y));
                }
                if (pts.Count > 1 && pts.First().GetDistanceTo(pts.Last()) < NearDupEps)
                    pts.RemoveAt(pts.Count - 1);
                result.Add(new Loop { Pts = pts, IsClosed = p3.Closed, Kind = "Polyline3d" });
            }
            else if (ent is Circle cir)
            {
                var pts = new List<Point2d>();
                for (int i = 0; i < CircleSides; i++)
                {
                    double ang = 2 * Math.PI * i / CircleSides;
                    pts.Add(new Point2d(
                        cir.Center.X + cir.Radius * Math.Cos(ang),
                        cir.Center.Y + cir.Radius * Math.Sin(ang)
                    ));
                }
                result.Add(new Loop { Pts = pts, IsClosed = true, Kind = "Circle" });
            }
            else if (ent is Arc arc)
            {
                var pts = new List<Point2d>();
                for (int i = 0; i <= ArcSides; i++)
                {
                    double ang = arc.StartAngle + (arc.EndAngle - arc.StartAngle) * i / ArcSides;
                    pts.Add(new Point2d(
                        arc.Center.X + arc.Radius * Math.Cos(ang),
                        arc.Center.Y + arc.Radius * Math.Sin(ang)
                    ));
                }
                result.Add(new Loop { Pts = pts, IsClosed = false, Kind = "Arc" });
            }
            else if (ent is Spline sp)
            {
                var pts = new List<Point2d>();
                double t0 = sp.StartParam, t1 = sp.EndParam;
                for (int i = 0; i <= SplineSamples; i++)
                {
                    double t = t0 + (t1 - t0) * i / SplineSamples;
                    var p = sp.GetPointAtParameter(t);
                    pts.Add(new Point2d(p.X, p.Y));
                }
                bool closed = pts.Count > 1 && pts.First().GetDistanceTo(pts.Last()) < NearDupEps;
                if (closed) pts.RemoveAt(pts.Count - 1);
                result.Add(new Loop { Pts = pts, IsClosed = closed, Kind = "Spline" });
            }

            return result;
        }

        // ---------- Validators ----------
        private List<Issue> ValidateLoop(Loop loop)
        {
            var raw = loop.Pts;
            var issues = new List<Issue>();

            if (raw == null || raw.Count < 3)
            {
                issues.Add(new Issue(raw.FirstOrDefault(), "Loop has < 3 vertices"));
                return issues;
            }

            if (!loop.IsClosed)
            {
                issues.Add(new Issue(raw[0], "Loop not closed"));
                return issues;
            }

            // 1) sanitize near-duplicates
            var pts = new List<Point2d>();
            for (int i = 0; i < raw.Count; i++)
            {
                var a = raw[i];
                var b = raw[(i + 1) % raw.Count];
                if (a.GetDistanceTo(b) < NearDupEps)
                    issues.Add(new Issue(b, "Near-duplicate vertex / zero-length edge"));
                else
                    pts.Add(a);
            }
            if (pts.Count < 3)
            {
                issues.Add(new Issue(raw[0], "Too few vertices after cleanup"));
                return issues;
            }

            // 2) nearly collinear
            for (int i = 0; i < pts.Count; i++)
            {
                var prev = pts[(i - 1 + pts.Count) % pts.Count];
                var curr = pts[i];
                var next = pts[(i + 1) % pts.Count];

                Vector2d v1 = (prev - curr);
                Vector2d v2 = (next - curr);

                if (v1.Length < Eps || v2.Length < Eps) continue;

                var ang = v1.GetAngleTo(v2) * (180.0 / Math.PI);
                if (Math.Abs(ang - 180.0) < CollinearDeg || ang < CollinearDeg)
                    issues.Add(new Issue(curr, "Nearly-collinear vertex"));
            }

            // 3) self-intersections
            for (int i = 0; i < pts.Count; i++)
            {
                var a1 = pts[i];
                var a2 = pts[(i + 1) % pts.Count];
                for (int j = i + 1; j < pts.Count; j++)
                {
                    if (Math.Abs(i - j) <= 1 || (i == 0 && j == pts.Count - 1)) continue;

                    var b1 = pts[j];
                    var b2 = pts[(j + 1) % pts.Count];

                    if (SegmentsIntersect(a1, a2, b1, b2, out Point2d hit, true))
                        issues.Add(new Issue(hit, "Self-intersection"));
                }
            }

            // 4) orientation
            if (SignedArea(pts) < 0)
                issues.Add(new Issue(Centroid(pts), "Outer loop is clockwise (should be CCW)"));

            return issues.Distinct(new IssueComparer()).ToList();
        }

        // ---------- Optional CDT probe (guarded) ----------
        private List<Issue> TryPoly2Tri(List<Point2d> loop)
        {
            var issues = new List<Issue>();
#if USE_POLY2TRI
            try
            {
                var contour = new List(Poly2Tri.PolygonPoint)();
                foreach (var p in loop)
                    contour.Add(new Poly2Tri.PolygonPoint(p.X, p.Y));
                var poly = new Poly2Tri.Polygon(contour);
                Poly2Tri.P2T.Triangulate(poly);
            }
            catch (System.Exception ex)
            {
                issues.Add(new Issue(Centroid(loop), "poly2tri failed: " + Short(ex.Message)));
            }
#endif
            return issues;
        }

        // ---------- Geometry helpers ----------
        private static double SignedArea(List<Point2d> pts)
        {
            double a = 0;
            for (int i = 0; i < pts.Count; i++)
            {
                var p = pts[i];
                var q = pts[(i + 1) % pts.Count];
                a += p.X * q.Y - q.X * p.Y;
            }
            return a * 0.5;
        }

        private static Point2d Centroid(List<Point2d> pts)
        {
            double cx = 0, cy = 0;
            foreach (var p in pts) { cx += p.X; cy += p.Y; }
            return new Point2d(cx / pts.Count, cy / pts.Count);
        }

        private static bool SegmentsIntersect(Point2d p1, Point2d p2, Point2d p3, Point2d p4, out Point2d hit, bool excludeEndpoints = false)
        {
            hit = new Point2d();
            double d = (p2.X - p1.X) * (p4.Y - p3.Y) - (p2.Y - p1.Y) * (p4.X - p3.X);
            if (Math.Abs(d) < 1e-12) return false;

            double ua = ((p4.X - p3.X) * (p1.Y - p3.Y) - (p4.Y - p3.Y) * (p1.X - p3.X)) / d;
            double ub = ((p2.X - p1.X) * (p1.Y - p3.Y) - (p2.Y - p1.Y) * (p1.X - p3.X)) / d;

            if (excludeEndpoints)
            {
                if (ua <= 1e-9 || ua >= 1 - 1e-9 || ub <= 1e-9 || ub >= 1 - 1e-9) return false;
            }

            if (ua >= 0 && ua <= 1 && ub >= 0 && ub <= 1)
            {
                hit = new Point2d(
                    p1.X + ua * (p2.X - p1.X),
                    p1.Y + ua * (p2.Y - p1.Y)
                );
                return true;
            }
            return false;
        }

        private static string Short(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            s = s.Replace("\r", " ").Replace("\n", " ");
            return s.Length > 120 ? s.Substring(0, 120) + "..." : s;
        }

        // ---------- Markers & reporting ----------
        private void DropMarkers(Transaction tr, Database db, List<Issue> issues)
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            foreach (var g in issues.GroupBy(i => i.Position))
            {
                var pos = g.Key;
                var label = string.Join("; ", g.Select(x => x.Message).Distinct());

                using (var c = new Circle(new Point3d(pos.X, pos.Y, 0), Vector3d.ZAxis, MarkerRadius))
                {
                    c.Layer = ErrorLayer;
                    c.Color = Color.FromRgb(255, 64, 64);
                    ms.AppendEntity(c);
                    tr.AddNewlyCreatedDBObject(c, true);
                }

                using (var t = new DBText())
                {
                    t.SetDatabaseDefaults();
                    t.Layer = ErrorLayer;
                    t.Color = Color.FromRgb(255, 64, 64);
                    t.TextString = label;
                    t.Height = TextHeight;
                    t.Position = new Point3d(pos.X + LabelOffset, pos.Y + LabelOffset, 0);
                    ms.AppendEntity(t);
                    tr.AddNewlyCreatedDBObject(t, true);
                }
            }
        }

        private void Report(Editor ed, string handle, string kind, List<Issue> issues)
        {
            ed.WriteMessage($"\nEntity {handle} ({kind}):");
            foreach (var i in issues)
                ed.WriteMessage($"\n  - {i.Message} @ ({i.Position.X:F3}, {i.Position.Y:F3})");
        }

        private void EnsureLayer(Database db, Transaction tr, string name, Color color)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(name))
            {
                lt.UpgradeOpen();
                var rec = new LayerTableRecord { Name = name, Color = color };
                lt.Add(rec);
                tr.AddNewlyCreatedDBObject(rec, true);
            }
        }
    }

    // =================================================================
    // ====== CURVE→LWP CONVERTER (Standard & Densify-Only modes) ======
    // =================================================================
    public class CurveToPolyline73
    {
        // ===== Defaults =====
        private const double DefaultMaxSeg = 50.0;
        private const double DefaultSag = 0.25;           // sampling error for curves
        private const double DefaultAngleCapDeg = 22.5;   // max angle step for curves
        private const double DefaultJoinTol = 1.0;        // base join tolerance
        private const bool AlwaysDeleteOriginals = true;

        // Cleanup & comparisons
        private const double NearDupEps = 1e-4;
        private const double CollinearDeg = 0.25;
        private const double DupTol = 1e-6;
        private const double LengthRelTol = 0.01;

        // ===== Commands =====
        [CommandMethod("CURVES2LWP", CommandFlags.Modal)]
        public void RunStd() => Execute(Mode.Standard);

        [CommandMethod("C2P", CommandFlags.Modal)]
        public void RunAlias() => Execute(Mode.Standard);

        // Densify-only for “cross” cases (no smoothing)
        [CommandMethod("C2PD", CommandFlags.Modal)]
        public void RunDensify() => Execute(Mode.DensifyOnly);

        private enum Mode { Standard, DensifyOnly }

        // ===== Core =====
        private void Execute(Mode mode)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // Select
            var pso = new PromptSelectionOptions { MessageForAdding = "\nSelect curves to convert: " };
            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "LINE,ARC,CIRCLE,ELLIPSE,LWPOLYLINE,POLYLINE,SPLINE")
            });
            var psr = ed.GetSelection(pso, filter);
            if (psr.Status != PromptStatus.OK) return;

            // Only prompt
            double maxSeg = PromptDoubleSafe(ed, "\nMax segment length", DefaultMaxSeg, 0.01, 1e9);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // Adaptive join tolerance
                Extents3d selExt = GetSelectionExtents(tr, psr.Value.GetObjectIds());
                double diag = DiagonalLength(selExt);
                double joinTol =
                    (mode == Mode.DensifyOnly)
                    ? Math.Max(DefaultJoinTol, Math.Min(0.03 * diag, 50.0))  // a bit looser for linework
                    : DefaultJoinTol;

                double sagTol = DefaultSag;
                double angleCap = DegreesToRadians(DefaultAngleCapDeg);

                var pieces = new List<Piece>();
                var toDelete = new HashSet<ObjectId>();

                // Gather & sample
                foreach (SelectedObject so in psr.Value)
                {
                    if (so == null) continue;
                    var ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                    if (!(ent is Curve cv)) continue;

                    toDelete.Add(ent.ObjectId);

                    var pts = TessellateCurveAdaptive(cv, maxSeg, sagTol, angleCap);
                    if (pts == null || pts.Count < 2) continue;

                    bool closedLike = IsClosedLike(cv, joinTol) || IsLoopByProximity(pts, joinTol);
                    var cleaned = CleanVertices(pts, closedLike, NearDupEps, CollinearDeg);
                    var st = Style.FromEntity(ent);

                    pieces.Add(new Piece
                    {
                        Points = cleaned,
                        IsClosed = closedLike,
                        Style = st,
                        SourceId = ent.ObjectId
                    });
                }

                // Join chains (both modes)
                pieces = SmartJoin(pieces, joinTol);

                // DENSIFY ONLY (no smoothing): insert intermediate points so each span <= maxSeg
                if (mode == Mode.DensifyOnly)
                {
                    for (int i = 0; i < pieces.Count; i++)
                    {
                        var p = pieces[i];
                        var dense = DensifyByMaxSeg(p.Points, p.IsClosed, maxSeg);
                        dense = CleanVertices(dense, p.IsClosed, NearDupEps, CollinearDeg);
                        p.Points = dense;
                        pieces[i] = p;
                    }
                }

                // Emit
                int created = 0, skippedDup = 0, skippedTwin = 0, skippedShort = 0;
                var emitted = new List<List<Point2d>>();
                double twinTol = Math.Max(joinTol, 4.0 * sagTol);

                foreach (var p in pieces)
                {
                    if (p.Points == null || p.Points.Count < 2) { skippedShort++; continue; }

                    var finalPts = p.Points;

                    if (p.IsClosed && finalPts.Count >= 3)
                    {
                        if (finalPts[0].GetDistanceTo(finalPts[finalPts.Count - 1]) <= DupTol)
                        {
                            var tmp = new List<Point2d>(finalPts);
                            if (tmp.Count > 1) tmp.RemoveAt(tmp.Count - 1);
                            finalPts = tmp;
                        }
                    }

                    finalPts = CleanVertices(finalPts, p.IsClosed, NearDupEps, CollinearDeg);

                    // Exact dup suppression (both modes)
                    if (IsDuplicate(emitted, finalPts, DupTol)) { skippedDup++; continue; }
                    // Near-twin only for Standard (NOT for DensifyOnly)
                    if (mode == Mode.Standard && IsNearTwin(emitted, finalPts, twinTol, LengthRelTol))
                    { skippedTwin++; continue; }

                    var plOut = new Polyline(finalPts.Count);
                    for (int i = 0; i < finalPts.Count; i++)
                        plOut.AddVertexAt(i, finalPts[i], 0.0, 0.0, 0.0);
                    plOut.Closed = p.IsClosed && finalPts.Count >= 3;
                    p.Style.ApplyTo(plOut);

                    btr.AppendEntity(plOut);
                    tr.AddNewlyCreatedDBObject(plOut, true);
                    created++;

                    emitted.Add(new List<Point2d>(finalPts));
                }

                // Erase originals
                if (AlwaysDeleteOriginals)
                {
                    foreach (var id in toDelete)
                    {
                        try
                        {
                            var src = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                            if (src != null && !src.IsErased) src.Erase();
                        }
                        catch { }
                    }
                }

                tr.Commit();
                ed.WriteMessage(
                    "\n{0}: {1} created. Skipped (dups:{2}, near-twins:{3}, too-short:{4}).",
                    (mode == Mode.DensifyOnly ? "C2PD" : "C2P"),
                    created, skippedDup, skippedTwin, skippedShort);
            }
        }

        // ===== Data structs =====
        private class Piece
        {
            public List<Point2d> Points;
            public bool IsClosed;
            public Style Style;
            public ObjectId SourceId;
        }

        private class Style
        {
            public string Layer;
            public short ColorIndex; // 256 = ByLayer
            public ObjectId LinetypeId;

            public static Style FromEntity(Entity e)
            {
                short aci = 256;
                if (e.Color != null)
                    aci = (short)(e.Color.IsByLayer ? 256 : e.Color.ColorIndex);
                return new Style { Layer = e.Layer, ColorIndex = aci, LinetypeId = e.LinetypeId };
            }

            public void ApplyTo(Entity e)
            {
                if (!string.IsNullOrEmpty(Layer)) e.Layer = Layer;
                e.Color = (ColorIndex == 256) ? Color.FromColorIndex(ColorMethod.ByLayer, 256)
                                              : Color.FromColorIndex(ColorMethod.ByAci, ColorIndex);
                if (LinetypeId.IsValid) e.LinetypeId = LinetypeId;
            }

            public string Key() { return (Layer ?? "") + "|" + ColorIndex.ToString() + "|" + LinetypeId.ToString(); }
        }

        // ===== Geometry & tessellation =====
        private static List<Point2d> TessellateCurveAdaptive(Curve cv, double maxSeg, double sagTol, double angleCap)
        {
            // Keep LWPOLYLINE vertices as-is (we may densify later, no smoothing)
            var lw = cv as Polyline;
            if (lw != null)
            {
                var ptsPL = new List<Point2d>();
                for (int i = 0; i < lw.NumberOfVertices; i++)
                    ptsPL.Add(lw.GetPoint2dAt(i));
                return ptsPL;
            }

            // Adaptively sample other curves (arc/ellipse/spline)
            var pts3 = SubdivideByError(cv, cv.StartParam, cv.EndParam, maxSeg, sagTol, angleCap);
            var pts2 = new List<Point2d>(pts3.Count);
            for (int i = 0; i < pts3.Count; i++)
                pts2.Add(new Point2d(pts3[i].X, pts3[i].Y));
            return pts2;
        }

        private static List<Point3d> SubdivideByError(Curve cv, double t0, double t1, double maxSeg, double sagTol, double angleCap)
        {
            var a = cv.GetPointAtParameter(t0);
            var b = cv.GetPointAtParameter(t1);

            bool tooLong = (a.DistanceTo(b) > maxSeg);

            double tm = 0.5 * (t0 + t1);
            var m = cv.GetPointAtParameter(tm);
            double err = DistancePointToSegment(m, a, b);

            bool angleExceeded = false;
            try
            {
                var ta = cv.GetFirstDerivative(t0);
                var tb = cv.GetFirstDerivative(t1);
                double ang = VectorAngle(ta, tb);
                angleExceeded = ang > angleCap;
            }
            catch { }

            if ((sagTol > 0 && err > sagTol) || tooLong || angleExceeded)
            {
                var left = SubdivideByError(cv, t0, tm, maxSeg, sagTol, angleCap);
                var right = SubdivideByError(cv, tm, t1, maxSeg, sagTol, angleCap);
                if (left.Count > 0) left.RemoveAt(left.Count - 1);
                left.AddRange(right);
                return left;
            }
            else
            {
                var outList = new List<Point3d>(2);
                outList.Add(a);
                outList.Add(b);
                return outList;
            }
        }

        private static bool IsClosedLike(Curve cv, double tol)
        {
            try { if (cv.Closed) return true; } catch { }
            var a = cv.StartPoint; var b = cv.EndPoint;
            return a.DistanceTo(b) <= tol;
        }

        private static bool IsLoopByProximity(List<Point2d> pts, double tol)
        {
            if (pts.Count < 3) return false;
            return pts[0].GetDistanceTo(pts[pts.Count - 1]) <= tol;
        }

        // ===== Densify (no smoothing) =====
        private static List<Point2d> DensifyByMaxSeg(List<Point2d> pts, bool closed, double maxSeg)
        {
            if (pts == null || pts.Count < 2) return pts;
            if (maxSeg <= 0) return pts;

            var outPts = new List<Point2d>();
            int last = pts.Count - 1;

            for (int i = 0; i < last; i++)
            {
                var p = pts[i];
                var q = pts[i + 1];
                outPts.Add(p);

                double dist = p.GetDistanceTo(q);
                if (dist > maxSeg)
                {
                    int divs = (int)Math.Floor(dist / maxSeg);
                    double step = 1.0 / (divs + 1);
                    for (int k = 1; k <= divs; k++)
                    {
                        double t = step * k;
                        var s = new Point2d(p.X + t * (q.X - p.X), p.Y + t * (q.Y - p.Y));
                        outPts.Add(s);
                    }
                }
            }
            outPts.Add(pts[last]);

            if (closed && outPts[0].GetDistanceTo(outPts[outPts.Count - 1]) > NearDupEps)
                outPts.Add(outPts[0]);

            return outPts;
        }

        // ===== Cleanup & helpers =====
        private static List<Point2d> CleanVertices(List<Point2d> raw, bool closed, double mergeTol, double collinearDeg)
        {
            if (raw == null || raw.Count == 0) return raw;

            // merge near duplicates
            var pts = new List<Point2d> { raw[0] };
            for (int i = 1; i < raw.Count; i++)
                if (raw[i].GetDistanceTo(pts[pts.Count - 1]) > mergeTol)
                    pts.Add(raw[i]);

            if (closed && pts.Count >= 2 && pts[0].GetDistanceTo(pts[pts.Count - 1]) <= mergeTol)
                pts[pts.Count - 1] = pts[0];

            // drop near-collinear
            if (pts.Count >= 3)
            {
                double cosThresh = Math.Cos(DegreesToRadians(180.0 - collinearDeg));
                var simp = new List<Point2d>();
                for (int i = 0; i < pts.Count; i++)
                {
                    if (!closed && (i == 0 || i == pts.Count - 1)) { simp.Add(pts[i]); continue; }

                    int prev = (i == 0) ? (closed ? pts.Count - 2 : 0) : i - 1;
                    int next = (i == pts.Count - 1) ? (closed ? 1 : pts.Count - 1) : i + 1;

                    var v1 = (pts[i] - pts[prev]);
                    var v2 = (pts[next] - pts[i]);

                    if (v1.Length <= 1e-12 || v2.Length <= 1e-12) { simp.Add(pts[i]); continue; }

                    v1 = v1.GetNormal();
                    v2 = v2.GetNormal();
                    double dot = v1.DotProduct(v2);
                    if (dot <= cosThresh) continue;

                    simp.Add(pts[i]);
                }
                pts = simp;

                if (closed && pts.Count >= 3 && pts[0].GetDistanceTo(pts[pts.Count - 1]) > mergeTol)
                    pts.Add(pts[0]);
            }
            return pts;
        }

        private static double DistancePointToSegment(Point3d p, Point3d a, Point3d b)
        {
            var ap = p - a; var ab = b - a;
            double ab2 = ab.DotProduct(ab);
            if (ab2 <= 1e-12) return ap.Length;
            double t = ap.DotProduct(ab) / ab2; t = Math.Max(0.0, Math.Min(1.0, t));
            var proj = a + t * ab; return (p - proj).Length;
        }

        private static double DegreesToRadians(double d) { return Math.PI * d / 180.0; }
        private static double VectorAngle(Vector3d a, Vector3d b)
        {
            if (a.Length < 1e-12 || b.Length < 1e-12) return 0.0;
            var na = a.GetNormal(); var nb = b.GetNormal();
            double v = Math.Max(-1.0, Math.Min(1.0, na.DotProduct(nb)));
            return Math.Acos(v);
        }

        // Extents
        private static Extents3d GetSelectionExtents(Transaction tr, ObjectId[] ids)
        {
            bool first = true;
            Extents3d ex = new Extents3d();
            for (int i = 0; i < ids.Length; i++)
            {
                try
                {
                    var e = tr.GetObject(ids[i], OpenMode.ForRead, false) as Entity;
                    if (e == null) continue;
                    var ee = e.GeometricExtents;
                    if (first) { ex = ee; first = false; }
                    else { ex.AddExtents(ee); }
                }
                catch { }
            }
            if (first) ex = new Extents3d(new Point3d(0, 0, 0), new Point3d(1, 1, 0));
            return ex;
        }
        private static double DiagonalLength(Extents3d ex)
        {
            try { return ex.MaxPoint.DistanceTo(ex.MinPoint); } catch { return 1.0; }
        }

        // ===== Duplicate / near-twin =====
        private static bool SamePath(List<Point2d> a, List<Point2d> b, double tol)
        {
            if (a == null || b == null) return false;
            if (a.Count != b.Count) return false;
            if (a.Count == 0) return true;

            bool fwd = true;
            for (int i = 0; i < a.Count; i++)
                if (a[i].GetDistanceTo(b[i]) > tol) { fwd = false; break; }
            if (fwd) return true;

            for (int i = 0; i < a.Count; i++)
                if (a[i].GetDistanceTo(b[b.Count - 1 - i]) > tol) return false;
            return true;
        }
        private static bool IsDuplicate(List<List<Point2d>> accum, List<Point2d> candidate, double tol)
        {
            for (int i = 0; i < accum.Count; i++)
                if (SamePath(accum[i], candidate, tol)) return true;
            return false;
        }

        private static double PathLength(List<Point2d> p)
        {
            double L = 0; for (int i = 1; i < p.Count; i++) L += p[i - 1].GetDistanceTo(p[i]); return L;
        }
        private static double MinDistPointToSeg(Point2d p, Point2d a, Point2d b)
        {
            var ap = p - a; var ab = b - a;
            double ab2 = ab.DotProduct(ab);
            if (ab2 <= 1e-12) return ap.Length;
            double t = Math.Max(0.0, Math.Min(1.0, ap.DotProduct(ab) / ab2));
            var proj = a + t * ab; return p.GetDistanceTo(proj);
        }
        private static double MinDistPointToPath(Point2d p, List<Point2d> path)
        {
            double best = double.MaxValue;
            for (int i = 1; i < path.Count; i++)
            {
                double d = MinDistPointToSeg(p, path[i - 1], path[i]);
                if (d < best) best = d;
            }
            return best;
        }
        private static double SymmetricHausdorff(List<Point2d> a, List<Point2d> b)
        {
            int stepA = Math.Max(1, a.Count / 64), stepB = Math.Max(1, b.Count / 64);
            double hAB = 0.0;
            for (int i = 0; i < a.Count; i += stepA) { double d = MinDistPointToPath(a[i], b); if (d > hAB) hAB = d; }
            double hBA = 0.0;
            for (int j = 0; j < b.Count; j += stepB) { double d = MinDistPointToPath(b[j], a); if (d > hBA) hBA = d; }
            return Math.Max(hAB, hBA);
        }
        private static bool IsNearTwin(List<List<Point2d>> accum, List<Point2d> candidate, double twinTol, double lengthRelTol)
        {
            double Lc = PathLength(candidate);
            for (int i = 0; i < accum.Count; i++)
            {
                var other = accum[i];
                double Lo = PathLength(other);
                double rel = Math.Abs(Lc - Lo) / Math.Max(1e-9, Math.Max(Lc, Lo));
                if (rel > lengthRelTol) continue;
                double h = SymmetricHausdorff(candidate, other);
                if (h <= twinTol) return true;
            }
            return false;
        }

        // ===== Smart-Join =====
        private static List<Piece> SmartJoin(List<Piece> pieces, double tol)
        {
            var closed = new List<Piece>();
            var open = new List<Piece>();
            foreach (var p in pieces) if (p.IsClosed) closed.Add(p); else open.Add(p);

            var result = new List<Piece>(); result.AddRange(closed);

            var byStyle = new Dictionary<string, List<Piece>>();
            foreach (var p in open)
            {
                var key = p.Style.Key();
                if (!byStyle.ContainsKey(key)) byStyle[key] = new List<Piece>();
                byStyle[key].Add(p);
            }

            foreach (var kv in byStyle)
            {
                var group = kv.Value;
                var used = new bool[group.Count];
                var chains = new List<List<Point2d>>();

                for (int i = 0; i < group.Count; i++)
                {
                    if (used[i]) continue;
                    var curr = new List<Point2d>(group[i].Points);
                    used[i] = true;

                    bool extended;
                    do
                    {
                        extended = false;
                        if (curr.Count > 0)
                        {
                            var tail = curr[curr.Count - 1];
                            if (TryFindAndConsume(ref curr, tail, group, used, tol, true))
                                extended = true;
                        }
                        if (curr.Count > 0)
                        {
                            var head = curr[0];
                            if (TryFindAndConsume(ref curr, head, group, used, tol, false))
                                extended = true;
                        }
                    }
                    while (extended);

                    bool makeClosed = curr.Count >= 3 && curr[0].GetDistanceTo(curr[curr.Count - 1]) <= tol;
                    chains.Add(CleanVertices(curr, makeClosed, NearDupEps, CollinearDeg));
                }

                foreach (var chain in chains)
                {
                    if (chain.Count < 2) continue;
                    bool isClosed = chain.Count >= 3 && chain[0].GetDistanceTo(chain[chain.Count - 1]) <= tol;
                    result.Add(new Piece { Points = chain, IsClosed = isClosed, Style = group[0].Style, SourceId = ObjectId.Null });
                }
            }
            return result;
        }

        private static bool TryFindAndConsume(ref List<Point2d> curr, Point2d anchor, List<Piece> group, bool[] used, double tol, bool atTail)
        {
            int bestIdx = -1; bool bestAtStart = false; double bestDist = tol;
            for (int i = 0; i < group.Count; i++)
            {
                if (used[i]) continue;
                var pts = group[i].Points; if (pts == null || pts.Count < 2) continue;
                double dStart = pts[0].GetDistanceTo(anchor); if (dStart <= bestDist) { bestDist = dStart; bestIdx = i; bestAtStart = true; }
                double dEnd = pts[pts.Count - 1].GetDistanceTo(anchor); if (dEnd <= bestDist) { bestDist = dEnd; bestIdx = i; bestAtStart = false; }
            }
            if (bestIdx < 0) return false;
            var add = new List<Point2d>(group[bestIdx].Points); if (!bestAtStart) add.Reverse();

            // avoid no-op joins (identical)
            if (SamePath(curr, add, tol)) { used[bestIdx] = true; return false; }

            if (atTail)
            {
                if (curr[curr.Count - 1].GetDistanceTo(add[0]) <= tol && add.Count > 0) add.RemoveAt(0);
                curr.AddRange(add);
            }
            else
            {
                if (curr[0].GetDistanceTo(add[add.Count - 1]) <= tol && add.Count > 0) add.RemoveAt(add.Count - 1);
                add.AddRange(curr); curr = add;
            }
            used[bestIdx] = true; return true;
        }

        // ===== UI helper =====
        private static double PromptDoubleSafe(Editor ed, string msg, double def, double min, double max)
        {
            var opts = new PromptDoubleOptions(msg + " <" + def + ">")
            {
                DefaultValue = def,
                UseDefaultValue = true,
                AllowNone = true
            };
            var res = ed.GetDouble(opts);
            double val = (res.Status == PromptStatus.OK) ? res.Value : def;
            if (val < min) val = min; if (val > max) val = max;
            return val;
        }
    }
}
