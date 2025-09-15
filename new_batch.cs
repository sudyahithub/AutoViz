// File: CurveToPolyline_CSharp73.cs
// Command: CURVES2LWP
// C# 7.3 SAFE VERSION (no index-from-end; no PromptDoubleOptions Lower/UpperLimit)
//
// Converts selected curves (LINE/ARC/CIRCLE/ELLIPSE/LW/2D/3D POLYLINE/SPLINE) to clean LWPOLYLINEs,
// ensures watertight closure for loops, optionally smart-joins open segments, and cleans vertices.

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

[assembly: CommandClass(typeof(LineAuditTool.CurveToPolyline73))]

namespace LineAuditTool
{
    public class CurveToPolyline73
    {
        // ===== Defaults =====
        private const double DefaultMaxSeg = 50.0;       // max chord length (DWG units)
        private const double DefaultSag = 0.25;          // sagitta/chord error
        private const double DefaultAngleCapDeg = 22.5;  // deg
        private const double DefaultJoinTol = 1.0;       // endpoint proximity for joins
        private const bool   DefaultDoJoin = true;
        private const bool   DefaultDeleteOriginals = false;

        // Cleaning
        private const double NearDupEps = 1e-4;
        private const double CollinearDeg = 0.25;

        [CommandMethod("CURVES2LWP", CommandFlags.Modal)]
        public void Run()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db  = doc.Database;
            var ed  = doc.Editor;

            // --- Selection ---
            var pso = new PromptSelectionOptions { MessageForAdding = "\nSelect curves to convert: " };
            var filter = new SelectionFilter(new[]
            {
                new TypedValue((int)DxfCode.Start, "LINE,ARC,CIRCLE,ELLIPSE,LWPOLYLINE,POLYLINE,SPLINE")
            });
            var psr = ed.GetSelection(pso, filter);
            if (psr.Status != PromptStatus.OK) return;

            // --- Options (validated manually; no LowerLimit/UpperLimit) ---
            double maxSeg   = PromptDoubleSafe(ed, "\nMax segment length", DefaultMaxSeg, 0.01, 1e9);
            double sagTol   = PromptDoubleSafe(ed, "\nSagitta (chord error) tolerance", DefaultSag, 0.0, 1e6);
            double angleCap = DegreesToRadians(PromptDoubleSafe(ed, "\nAngle cap per step (deg)", DefaultAngleCapDeg, 0.1, 180.0));
            double joinTol  = PromptDoubleSafe(ed, "\nJoin tolerance (endpoint proximity)", DefaultJoinTol, 0.0, 1e6);
            bool doJoin     = PromptYesNo(ed, "\nSmart-join open segments", DefaultDoJoin);
            bool deleteOriginals = PromptYesNo(ed, "\nDelete originals after conversion", DefaultDeleteOriginals);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                var pieces = new List<Piece>();

                foreach (SelectedObject so in psr.Value)
                {
                    if (so == null) continue;
                    var ent = tr.GetObject(so.ObjectId, deleteOriginals ? OpenMode.ForWrite : OpenMode.ForRead) as Entity;
                    if (!(ent is Curve cv)) continue;

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

                // Optional join
                if (doJoin)
                    pieces = SmartJoin(pieces, joinTol);

                // Emit
                int created = 0;
                foreach (var p in pieces)
                {
                    if (p.Points.Count < 2) continue;

                    var finalPts = p.Points;

                    // Force watertight if closed
                    if (p.IsClosed && finalPts.Count >= 3)
                    {
                        if (finalPts[0].GetDistanceTo(finalPts[finalPts.Count - 1]) > joinTol)
                        {
                            var tmp = new List<Point2d>(finalPts);
                            tmp.Add(finalPts[0]);
                            finalPts = tmp;
                        }
                    }

                    finalPts = CleanVertices(finalPts, p.IsClosed, NearDupEps, CollinearDeg);

                    var pl = new Polyline(finalPts.Count);
                    for (int i = 0; i < finalPts.Count; i++)
                        pl.AddVertexAt(i, finalPts[i], 0.0, 0.0, 0.0);
                    pl.Closed = p.IsClosed && finalPts.Count >= 3;

                    p.Style.ApplyTo(pl);

                    btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);
                    created++;

                    if (deleteOriginals && p.SourceId.IsValid && !p.SourceId.IsNull)
                    {
                        try
                        {
                            var src = (Entity)tr.GetObject(p.SourceId, OpenMode.ForWrite, false);
                            if (src != null) src.Erase();
                        }
                        catch { /* ignore */ }
                    }
                }

                tr.Commit();
                ed.WriteMessage("\nCURVES2LWP: Created {0} LWPOLYLINE(s).", created);
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

                return new Style
                {
                    Layer = e.Layer,
                    ColorIndex = aci,
                    LinetypeId = e.LinetypeId
                };
            }

            public void ApplyTo(Entity e)
            {
                if (!string.IsNullOrEmpty(Layer)) e.Layer = Layer;
                e.Color = (ColorIndex == 256) ? Color.FromColorIndex(ColorMethod.ByLayer, 256)
                                              : Color.FromColorIndex(ColorMethod.ByAci, ColorIndex);
                if (LinetypeId.IsValid) e.LinetypeId = LinetypeId;
            }

            public string Key()
            {
                return (Layer ?? "") + "|" + ColorIndex.ToString() + "|" + LinetypeId.ToString();
            }
        }

        // ===== Geometry & tessellation =====
        private static List<Point2d> TessellateCurveAdaptive(Curve cv, double maxSeg, double sagTol, double angleCap)
        {
            // Already linear?
            var pl = cv as Polyline;
            if (pl != null)
            {
                var ptsPL = new List<Point2d>();
                for (int i = 0; i < pl.NumberOfVertices; i++)
                    ptsPL.Add(pl.GetPoint2dAt(i));
                return ptsPL;
            }

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
                if (left.Count > 0) left.RemoveAt(left.Count - 1); // drop duplicate mid
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
            var a = cv.StartPoint;
            var b = cv.EndPoint;
            return a.DistanceTo(b) <= tol;
        }

        private static bool IsLoopByProximity(List<Point2d> pts, double tol)
        {
            if (pts.Count < 3) return false;
            return pts[0].GetDistanceTo(pts[pts.Count - 1]) <= tol;
        }

        private static List<Point2d> CleanVertices(List<Point2d> raw, bool closed, double mergeTol, double collinearDeg)
        {
            if (raw == null || raw.Count == 0) return raw;

            // 1) merge near-duplicates
            var pts = new List<Point2d>();
            pts.Add(raw[0]);
            for (int i = 1; i < raw.Count; i++)
            {
                if (raw[i].GetDistanceTo(pts[pts.Count - 1]) > mergeTol)
                    pts.Add(raw[i]);
            }
            if (closed && pts.Count >= 2 && pts[0].GetDistanceTo(pts[pts.Count - 1]) <= mergeTol)
                pts[pts.Count - 1] = pts[0];

            // 2) drop collinear points using Vector2d math (C# 7.3-safe)
            if (pts.Count >= 3)
            {
                double cosThresh = Math.Cos(DegreesToRadians(180.0 - collinearDeg)); // ~ -0.99999
                var simp = new List<Point2d>();
                for (int i = 0; i < pts.Count; i++)
                {
                    if (!closed && (i == 0 || i == pts.Count - 1))
                    {
                        simp.Add(pts[i]);
                        continue;
                    }

                    int prev = (i == 0) ? (closed ? pts.Count - 2 : 0) : i - 1;
                    int next = (i == pts.Count - 1) ? (closed ? 1 : pts.Count - 1) : i + 1;

                    var v1 = (pts[i] - pts[prev]); // Vector2d
                    var v2 = (pts[next] - pts[i]); // Vector2d

                    if (v1.Length <= 1e-12 || v2.Length <= 1e-12)
                    {
                        simp.Add(pts[i]);
                        continue;
                    }

                    v1 = v1.GetNormal();
                    v2 = v2.GetNormal();

                    double dot = v1.DotProduct(v2);

                    // If dot ≈ -1 (straight line), drop middle
                    if (dot <= cosThresh)
                        continue;

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
            var ap = p - a;
            var ab = b - a;
            double ab2 = ab.DotProduct(ab);
            if (ab2 <= 1e-12) return ap.Length;
            double t = ap.DotProduct(ab) / ab2;
            t = Math.Max(0.0, Math.Min(1.0, t));
            var proj = a + t * ab;
            return (p - proj).Length;
        }

        private static double VectorAngle(Vector3d a, Vector3d b)
        {
            if (a.Length < 1e-12 || b.Length < 1e-12) return 0.0;
            var na = a.GetNormal();
            var nb = b.GetNormal();
            double d = Math.Max(-1.0, Math.Min(1.0, na.DotProduct(nb)));
            return Math.Acos(d);
        }

        private static double DegreesToRadians(double d) { return Math.PI * d / 180.0; }

        // ===== Smart Join (C# 7.3-safe) =====
        private class Endpoint
        {
            public Point2d P;
            public int PieceIndex;
            public bool IsStart;
            public string StyleKey;
        }

        private static List<Piece> SmartJoin(List<Piece> pieces, double tol)
        {
            var closed = new List<Piece>();
            var open = new List<Piece>();
            for (int i = 0; i < pieces.Count; i++)
            {
                if (pieces[i].IsClosed) closed.Add(pieces[i]); else open.Add(pieces[i]);
            }

            var result = new List<Piece>();
            result.AddRange(closed);

            // Group by style
            var byStyle = new Dictionary<string, List<Piece>>();
            for (int i = 0; i < open.Count; i++)
            {
                var key = open[i].Style.Key();
                List<Piece> list;
                if (!byStyle.TryGetValue(key, out list))
                {
                    list = new List<Piece>();
                    byStyle[key] = list;
                }
                list.Add(open[i]);
            }

            foreach (var kv in byStyle)
            {
                var styleKey = kv.Key;
                var group = kv.Value;

                var endpoints = new List<Endpoint>();
                for (int i = 0; i < group.Count; i++)
                {
                    var pts = group[i].Points;
                    if (pts == null || pts.Count < 2) continue;

                    endpoints.Add(new Endpoint { P = pts[0], StyleKey = styleKey, PieceIndex = i, IsStart = true });
                    endpoints.Add(new Endpoint { P = pts[pts.Count - 1], StyleKey = styleKey, PieceIndex = i, IsStart = false });
                }

                // Buckets by tol
                var buckets = new Dictionary<string, List<Endpoint>>();
                for (int i = 0; i < endpoints.Count; i++)
                {
                    var k = Hash(endpoints[i].P, tol);
                    List<Endpoint> lst;
                    if (!buckets.TryGetValue(k, out lst))
                    {
                        lst = new List<Endpoint>();
                        buckets[k] = lst;
                    }
                    lst.Add(endpoints[i]);
                }

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

                        // tail
                        if (curr.Count > 0)
                        {
                            var tail = curr[curr.Count - 1];
                            if (TryFindAndConsume(ref curr, tail, group, used, buckets, tol, true))
                                extended = true;
                        }
                        // head
                        if (curr.Count > 0)
                        {
                            var head = curr[0];
                            if (TryFindAndConsume(ref curr, head, group, used, buckets, tol, false))
                                extended = true;
                        }
                    }
                    while (extended);

                    bool makeClosed = curr.Count >= 3 && curr[0].GetDistanceTo(curr[curr.Count - 1]) <= tol;
                    chains.Add(CleanVertices(curr, makeClosed, NearDupEps, CollinearDeg));
                }

                // Emit
                for (int c = 0; c < chains.Count; c++)
                {
                    var chain = chains[c];
                    if (chain.Count < 2) continue;
                    bool isClosed = chain.Count >= 3 && chain[0].GetDistanceTo(chain[chain.Count - 1]) <= tol;
                    result.Add(new Piece
                    {
                        Points = chain,
                        IsClosed = isClosed,
                        Style = group[0].Style,
                        SourceId = ObjectId.Null
                    });
                }
            }

            return result;
        }

        private static string Hash(Point2d p, double tol)
        {
            long x = (long)Math.Round(p.X / tol);
            long y = (long)Math.Round(p.Y / tol);
            return x.ToString() + ":" + y.ToString();
        }

        private static bool TryFindAndConsume(ref List<Point2d> curr, Point2d anchor, List<Piece> group, bool[] used,
                                              Dictionary<string, List<Endpoint>> buckets, double tol, bool atTail)
        {
            List<Endpoint> list;
            if (!buckets.TryGetValue(Hash(anchor, tol), out list)) return false;

            for (int j = 0; j < list.Count; j++)
            {
                var ep = list[j];
                if (used[ep.PieceIndex]) continue;

                var pts = group[ep.PieceIndex].Points;
                if (pts == null || pts.Count < 2) continue;

                double startDist = pts[0].GetDistanceTo(anchor);
                double endDist   = pts[pts.Count - 1].GetDistanceTo(anchor);

                if (startDist <= tol || endDist <= tol)
                {
                    var add = new List<Point2d>(pts);
                    if (endDist <= tol)
                        add.Reverse(); // connect correctly

                    if (atTail)
                    {
                        if (curr[curr.Count - 1].GetDistanceTo(add[0]) <= tol && add.Count > 0)
                            add.RemoveAt(0);
                        curr.AddRange(add);
                    }
                    else
                    {
                        if (curr[0].GetDistanceTo(add[add.Count - 1]) <= tol && add.Count > 0)
                            add.RemoveAt(add.Count - 1);
                        add.AddRange(curr);
                        curr = add;
                    }

                    used[ep.PieceIndex] = true;
                    return true;
                }
            }
            return false;
        }

        // ===== UI helpers (no LowerLimit/UpperLimit) =====
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
            if (double.IsNaN(val) || double.IsInfinity(val)) val = def;
            if (val < min) val = min;
            if (val > max) val = max;
            return val;
        }

        private static bool PromptYesNo(Editor ed, string msg, bool def)
        {
            var p = new PromptKeywordOptions(msg + " (" + (def ? "Yes/No" : "No/Yes") + ")")
            {
                AllowArbitraryInput = false
            };
            p.Keywords.Add("Yes");
            p.Keywords.Add("No");
            p.Keywords.Default = def ? "Yes" : "No";
            var r = ed.GetKeywords(p);
            return (r.Status == PromptStatus.OK) ? (r.StringResult == "Yes") : def;
        }
    }
}
