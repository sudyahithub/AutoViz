// GeometrySignature.cs
// Extract lightweight vector features and a 0..1 similarity.
// C# 7.3.

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;

namespace LineAuditTool.Matching
{
    internal static class GeometrySignature
    {
        // Produces GeomFeatures from a block definition (definition space).
        public static GeomFeatures Compute(BlockTableRecord btr, Transaction tr)
        {
            GeomFeatures gf = new GeomFeatures();
            gf.AngleHist8 = new int[8];

            // Extents
            Extents3d? extOpt = null;
            foreach (ObjectId id in btr)
            {
                Entity e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (e == null || Skip(e)) continue;
                try
                {
                    Extents3d ex = e.GeometricExtents;
                    if (!extOpt.HasValue) extOpt = ex;
                    else
                    {
                        Extents3d cur = extOpt.Value;
                        Point3d min = new Point3d(Math.Min(cur.MinPoint.X, ex.MinPoint.X),
                                                  Math.Min(cur.MinPoint.Y, ex.MinPoint.Y), 0);
                        Point3d max = new Point3d(Math.Max(cur.MaxPoint.X, ex.MaxPoint.X),
                                                  Math.Max(cur.MaxPoint.Y, ex.MaxPoint.Y), 0);
                        extOpt = new Extents3d(min, max);
                    }
                }
                catch { }
            }
            if (extOpt.HasValue)
            {
                Extents3d ext = extOpt.Value;
                double w = ext.MaxPoint.X - ext.MinPoint.X;
                double d = ext.MaxPoint.Y - ext.MinPoint.Y;
                if (w < 1e-6) w = 1e-6;
                if (d < 1e-6) d = 1e-6;
                gf.WidthMM = w;
                gf.DepthMM = d;
                gf.Aspect = w / d;
            }

            // Counts, angle histogram and components via bbox proximity
            List<Extents3d> boxes = new List<Extents3d>(64);

            foreach (ObjectId id in btr)
            {
                Entity e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (e == null || Skip(e)) continue;

                if (e is Line) gf.LineCount++;
                else if (e is Arc) gf.ArcCount++;
                else if (e is Spline) gf.SplineCount++;
                else if (e is Polyline) gf.PolyCount++;

                try
                {
                    boxes.Add(e.GeometricExtents);
                }
                catch { }

                // Angle contributions
                if (e is Line l)
                {
                    AddAngle(gf.AngleHist8, AngleOf(l.StartPoint, l.EndPoint));
                }
                else if (e is Polyline pl)
                {
                    int n = pl.NumberOfVertices;
                    int i;
                    for (i = 0; i < n - 1; i++)
                    {
                        Point2d p1 = pl.GetPoint2dAt(i);
                        Point2d p2 = pl.GetPoint2dAt(i + 1);
                        AddAngle(gf.AngleHist8, AngleOf(new Point3d(p1.X, p1.Y, 0), new Point3d(p2.X, p2.Y, 0)));
                    }
                }
                else if (e is Arc ar)
                {
                    double mid = (ar.StartAngle + ar.EndAngle) * 0.5;
                    // tangent angle (convert to segment orientation)
                    double ang = mid + Math.PI * 0.5;
                    while (ang < 0) ang += Math.PI;
                    while (ang >= Math.PI) ang -= Math.PI;
                    int bin = (int)(8.0 * ang / Math.PI);
                    if (bin < 0) bin = 0; if (bin > 7) bin = 7;
                    gf.AngleHist8[bin]++;
                }
            }

            // Normalize histogram to sum 1000 (integer) to avoid doubles
            int sum = 0; int iBin;
            for (iBin = 0; iBin < 8; iBin++) sum += gf.AngleHist8[iBin];
            if (sum > 0)
            {
                for (iBin = 0; iBin < 8; iBin++)
                    gf.AngleHist8[iBin] = (int)Math.Round(1000.0 * gf.AngleHist8[iBin] / sum);
            }

            gf.Components = EstimateComponents(boxes, Math.Max(gf.WidthMM, gf.DepthMM) * 0.01);

            return gf;
        }

        private static bool Skip(Entity e)
        {
            if (e is DBText || e is MText) return true;
            if (e is Hatch) return true;
            if (e is Dimension) return true;
            if (e is Leader || e is MLeader) return true;
            return false;
        }

        private static double AngleOf(Point3d a, Point3d b)
        {
            double dx = b.X - a.X;
            double dy = b.Y - a.Y;
            double ang = Math.Atan2(Math.Abs(dy), Math.Abs(dx)); // 0..pi/2; fold by symmetry
            // map to 0..pi
            double phi = Math.Atan2(dy, dx);
            while (phi < 0) phi += Math.PI;
            while (phi >= Math.PI) phi -= Math.PI;
            return phi;
        }

        private static void AddAngle(int[] hist8, double phi)
        {
            int bin = (int)(8.0 * phi / Math.PI);
            if (bin < 0) bin = 0; if (bin > 7) bin = 7;
            hist8[bin]++;
        }

        private static int EstimateComponents(List<Extents3d> boxes, double thr)
        {
            int n = boxes.Count;
            if (n == 0) return 0;
            int[] parent = new int[n];
            int i;
            for (i = 0; i < n; i++) parent[i] = i;

            Func<int, int> find = null;
            find = (int x) =>
            {
                while (parent[x] != x) { parent[x] = parent[parent[x]]; x = parent[x]; }
                return x;
            };

            Action<int, int> unite = (a, b) =>
            {
                int ra = find(a), rb = find(b);
                if (ra != rb) parent[rb] = ra;
            };

            int a, b;
            for (a = 0; a < n; a++)
            {
                for (b = a + 1; b < n; b++)
                {
                    if (Close(boxes[a], boxes[b], thr))
                        unite(a, b);
                }
            }

            int comps = 0;
            for (i = 0; i < n; i++)
                if (find(i) == i) comps++;
            return comps;
        }

        private static bool Close(Extents3d A, Extents3d B, double thr)
        {
            double dx = Distance1D(A.MinPoint.X, A.MaxPoint.X, B.MinPoint.X, B.MaxPoint.X);
            double dy = Distance1D(A.MinPoint.Y, A.MaxPoint.Y, B.MinPoint.Y, B.MaxPoint.Y);
            return (dx <= thr && dy <= thr);
        }

        private static double Distance1D(double a0, double a1, double b0, double b1)
        {
            double left = Math.Max(a0, b0);
            double right = Math.Min(a1, b1);
            if (right >= left) return 0.0;           // overlap
            return left - right;                      // gap
        }

        // 0..1 similarity
        public static double GeomSimilarity(ref GeomFeatures A, ref GeomFeatures B)
        {
            // Sizes (width & aspect)
            double Wa = A.WidthMM, Wb = B.WidthMM;
            double Da = A.DepthMM, Db = B.DepthMM;
            if (Wa < 1e-6) Wa = 1e-6; if (Wb < 1e-6) Wb = 1e-6;
            if (Da < 1e-6) Da = 1e-6; if (Db < 1e-6) Db = 1e-6;

            double sW = Math.Exp(-Math.Abs(Wa - Wb) / Math.Max(Wa, Wb));
            double sD = Math.Exp(-Math.Abs(Da - Db) / Math.Max(Da, Db));
            double sAspect = Math.Exp(-Math.Abs(Math.Log((A.Aspect + 1e-9) / (B.Aspect + 1e-9))));

            // Angle histogram chi-square
            double chi = 0.0;
            int i;
            for (i = 0; i < 8; i++)
            {
                double x = A.AngleHist8[i];
                double y = B.AngleHist8[i];
                double denom = x + y;
                if (denom > 0) chi += ((x - y) * (x - y)) / denom;
            }
            // map chi to sim (smaller chi -> sim close to 1)
            double sAng = Math.Exp(-chi / 1000.0);

            // Counts difference normalized
            double cA = A.LineCount + A.ArcCount + A.SplineCount + A.PolyCount;
            double cB = B.LineCount + B.ArcCount + B.SplineCount + B.PolyCount;
            if (cA < 1) cA = 1; if (cB < 1) cB = 1;
            double sCnt = Math.Exp(-Math.Abs(cA - cB) / Math.Max(cA, cB));

            // Components diff
            int compA = A.Components <= 0 ? 1 : A.Components;
            int compB = B.Components <= 0 ? 1 : B.Components;
            double sComp = Math.Exp(-Math.Abs(compA - compB) / (double)Math.Max(compA, compB));

            double geom = 0.30 * sW + 0.20 * sD + 0.20 * sAspect + 0.15 * sAng + 0.10 * sCnt + 0.05 * sComp;
            if (geom < 0.0) geom = 0.0; if (geom > 1.0) geom = 1.0;
            return geom;
        }
    }
}
