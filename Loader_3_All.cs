// File: LineAuditCommand.cs
// Commands:
//   POLY2TRI_VALIDATE       → find & mark CDT-breaking issues
//   POLY2TRI_CLEAR_ERRORS   → delete all markers/labels on _TRI_ERRORS
//   BATCH_ARC2LIN           → tessellate arcs/splines/ellipses/polylines to LWPolyline
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
[assembly: CommandClass(typeof(LineAuditTool.Batch))]

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
        private const double TextHeight = 2.2;   // fixed height
        private const double LabelOffset = 8.0;   // XY offset from marker center

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
    // ============= ARC/SPLINE/ELLIPSE → POLYLINE CONVERTER ===========
    // =================================================================
    public class Batch
    {
        // --- Config ---
        const double MaxSeg = 3.0;           // max chord length in DWG units
        const bool DeleteOriginals = true;   // erase sources after conversion
        const bool DoJoin = true;            // join touching polylines after
        const double AngleCap = Math.PI / 6.0; // max angle per step (30°)

        [CommandMethod("BATCH_ARC2LIN")]
        public void Run()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var pso = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect arcs/splines/ellipses/polylines to convert (Enter to finish):"
            };
            var psr = ed.GetSelection(pso);
            if (psr.Status != PromptStatus.OK) return;

            var created = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                foreach (SelectedObject so in psr.Value)
                {
                    if (so == null) continue;
                    var ent = tr.GetObject(so.ObjectId, OpenMode.ForWrite) as Entity;
                    if (ent == null) continue;

                    var pl = TessellateToPolyline(ent, MaxSeg);
                    if (pl == null) continue;

                    var id = btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);
                    created.Add(id);

                    if (DeleteOriginals) ent.Erase();
                }

                if (DoJoin && created.Count > 1)
                    JoinPolylines(created, tr, passes: 2);

                tr.Commit();
            }

            if (created.Count > 0)
            {
                ed.SetImpliedSelection(created.ToArray());
                ed.WriteMessage($"\nConverted/affected: {created.Count}");
            }
            else
            {
                ed.WriteMessage("\nNo convertible entities found.");
            }
        }

        // =========================================================
        // Tessellation: length-based + adaptive refinement
        // =========================================================
        private static Polyline TessellateToPolyline(Entity e, double maxSeg)
        {
            if (e is Curve c)
            {
                var pts = SampleCurveByLengthAdaptive(c, maxSeg, AngleCap);
                if (pts == null || pts.Count < 2) return null;

                bool closed = false;
                try { closed = c.Closed; } catch { }
                double elev = pts.Count > 0 ? pts[0].Z : 0.0;
                return MakePolyline(pts, closed, elev);
            }

            switch (e)
            {
                case Polyline2d p2: return ConvertPolyline2d(p2);
                case Polyline3d p3: return ConvertPolyline3d(p3);
                default: return null;
            }
        }

        private static List<Point3d> SampleCurveByLengthAdaptive(Curve c, double maxSeg, double maxAngleStep)
        {
            var pts = new List<Point3d>();

            double len;
            try { len = c.GetDistanceAtParameter(c.EndParam); }
            catch { len = EstimateLength(c); }

            if (len <= 0) return pts;

            int steps = Math.Max(2, (int)Math.Ceiling(len / Math.Max(1e-9, maxSeg)));
            for (int i = 0; i <= steps; i++)
            {
                double d = Math.Min(len, i * (len / steps));
                pts.Add(SafePointAtDist(c, d));
            }

            double tol = Math.Max(1e-6, maxSeg * 0.125);
            pts = RefineByDeviation(c, pts, tol, maxAngleStep);

            return pts;
        }

        private static List<Point3d> RefineByDeviation(Curve c, List<Point3d> pts, double tol, double maxAngleStep)
        {
            var outPts = new List<Point3d>();
            for (int i = 0; i < pts.Count - 1; i++)
            {
                var a = pts[i];
                var b = pts[i + 1];
                outPts.Add(a);

                double da = SafeDistAtPoint(c, a);
                double db = SafeDistAtPoint(c, b);
                double dm = (da + db) * 0.5;
                var mCurve = SafePointAtDist(c, dm);

                double dev = DistancePointToSegment(mCurve, a, b);

                bool angleTooBig = false;
                if (maxAngleStep > 0)
                {
                    Vector3d va = (mCurve - a);   // Already Vector3d
                    Vector3d vb = (b - mCurve);
                    double ang = AngleBetween(va, vb);
                    angleTooBig = ang > maxAngleStep;
                }

                if (dev > tol || angleTooBig)
                {
                    outPts.Add(mCurve);
                }
            }
            outPts.Add(pts[pts.Count - 1]);

            if (outPts.Count > pts.Count)
                return RefineByDeviation(c, outPts, tol, maxAngleStep);

            return outPts;
        }

        private static Point3d SafePointAtDist(Curve c, double dist)
        {
            try { return c.GetPointAtDist(dist); }
            catch
            {
                double t0 = c.StartParam, t1 = c.EndParam;
                return c.GetPointAtParameter(t0 + (t1 - t0) * 0.5);
            }
        }
        private static double SafeDistAtPoint(Curve c, Point3d p)
        {
            try { return c.GetDistAtPoint(p); }
            catch
            {
                double t0 = c.StartParam, t1 = c.EndParam;
                return 0.5 * (c.GetDistanceAtParameter(t0) + c.GetDistanceAtParameter(t1));
            }
        }
        private static double EstimateLength(Curve c)
        {
            int n = 32;
            double t0 = c.StartParam, t1 = c.EndParam;
            var last = c.GetPointAtParameter(t0);
            double sum = 0;
            for (int i = 1; i <= n; i++)
            {
                double t = t0 + (t1 - t0) * (i / (double)n);
                var p = c.GetPointAtParameter(t);
                sum += last.DistanceTo(p);
                last = p;
            }
            return sum;
        }

        private static double DistancePointToSegment(Point3d p, Point3d a, Point3d b)
        {
            var ap = p - a;
            var ab = b - a;
            double t = (ap.DotProduct(ab)) / Math.Max(1e-12, ab.DotProduct(ab));
            t = Math.Max(0, Math.Min(1, t));
            var proj = a + ab * t;
            return p.DistanceTo(proj);
        }
        private static double AngleBetween(Vector3d a, Vector3d b)
        {
            double d = a.DotProduct(b) / (a.Length * b.Length + 1e-12);
            d = Math.Max(-1.0, Math.Min(1.0, d));
            return Math.Acos(d);
        }

        private static Polyline MakePolyline(IList<Point3d> pts, bool closed, double elevation)
        {
            var pl = new Polyline(pts.Count);
            for (int i = 0; i < pts.Count; i++)
                pl.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), 0.0, 0.0, 0.0);
            pl.Closed = closed;
            pl.Elevation = elevation;
            pl.SetDatabaseDefaults();
            return pl;
        }
        private static Polyline ConvertPolyline2d(Polyline2d p2)
        {
            var pts = new List<Point3d>();
            var tm = p2.Database.TransactionManager;
            foreach (ObjectId vId in p2)
            {
                var v = (Vertex2d)tm.GetObject(vId, OpenMode.ForRead);
                pts.Add(v.Position);
            }
            return MakePolyline(pts, p2.Closed, pts.Count > 0 ? pts[0].Z : 0.0);
        }
        private static Polyline ConvertPolyline3d(Polyline3d p3)
        {
            var pts = new List<Point3d>();
            var tm = p3.Database.TransactionManager;
            foreach (ObjectId vId in p3)
            {
                var v = (PolylineVertex3d)tm.GetObject(vId, OpenMode.ForRead);
                pts.Add(v.Position);
            }
            return MakePolyline(pts, p3.Closed, pts.Count > 0 ? pts[0].Z : 0.0);
        }

        private static void JoinPolylines(List<ObjectId> ids, Transaction tr, int passes = 1)
        {
            for (int pass = 0; pass < passes; pass++)
            {
                var survivors = new List<ObjectId>();
                var pool = new HashSet<ObjectId>(ids);

                while (pool.Count > 0)
                {
                    var first = pool.First();
                    pool.Remove(first);

                    var pl = tr.GetObject(first, OpenMode.ForWrite) as Polyline;
                    if (pl == null) { survivors.Add(first); continue; }

                    var joined = new List<ObjectId>();
                    foreach (var other in pool.ToList())
                    {
                        var cur = tr.GetObject(other, OpenMode.ForWrite) as Entity;
                        if (cur == null || cur.ObjectId == pl.ObjectId) continue;

                        var arr = new Entity[] { cur };
                        int before = pl.NumberOfVertices;
                        pl.JoinEntities(arr);
                        if (pl.NumberOfVertices != before)
                            joined.Add(other);
                    }

                    foreach (var j in joined)
                    {
                        pool.Remove(j);
                        var e = (Entity)tr.GetObject(j, OpenMode.ForWrite);
                        e.Erase();
                    }
                    survivors.Add(first);
                }

                ids.Clear();
                ids.AddRange(survivors);
            }
        }
    }
}
