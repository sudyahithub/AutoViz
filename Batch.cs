using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

[assembly: CommandClass(typeof(LineAuditTool.Batch))]

namespace LineAuditTool
{
    public class Batch
    {
        // --- Config ---
        const double MaxSeg = 3.0;           // max chord length in DWG units
        const bool DeleteOriginals = true; // erase sources after conversion
        const bool DoJoin = true; // join touching polylines after
        const double AngleCap = Math.PI / 6.0; // optional angle limiter per step (30°)

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
                    // FIX: (p - p) already returns Vector3d; no GetAsVector()
                    Vector3d va = (mCurve - a);
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



