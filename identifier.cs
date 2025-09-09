// File: LineAuditCommand.cs
// Commands:
//   POLY2TRI_VALIDATE     → find & mark issues
//   POLY2TRI_CLEAR_ERRORS → delete all markers/labels on _TRI_ERRORS
//
// Finds CDT-breaking issues directly in AutoCAD.
// Handles LWPOLYLINE/2D/3D (open/closed), CIRCLE, ARC, SPLINE (tessellated).
// Drops red circles + DBText labels on layer _TRI_ERRORS.

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

[assembly: CommandClass(typeof(LineAuditTool.BlockMatcher))]

namespace LineAuditTool
{
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
        private const double TextHeight = 2.2;   // fixed height (was 0.4 * radius ≈ 8)
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
                var contour = new List<Poly2Tri.PolygonPoint>();
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
}
