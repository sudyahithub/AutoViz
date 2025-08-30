// PreviewRenderer.cs
// Deterministic 256x256 preview pipeline with mirror-correction and thinning.
// C# 7.3; avoids LINQ in hot loops.

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace LineAuditTool.Matching
{
    public sealed class PreviewRenderer
    {
        // 8% padding around extents
        private const double PadFactor = 0.08;

        // Cache: key = (BlockTableRecord.ObjectId, parityOdd)
        private readonly Dictionary<ulong, BlockPreview> _cache = new Dictionary<ulong, BlockPreview>();

        public BlockPreview GetPreview(BlockReference br, Transaction tr)
        {
            // Cache key: BTR id + parity (exactly one negative axis)
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
            bool parityOdd = (br.ScaleFactors.X < 0) ^ (br.ScaleFactors.Y < 0);
            ulong key = (((ulong)btr.ObjectId.OldIdPtr.ToInt64()) << 1) | (parityOdd ? 1UL : 0UL);

            BlockPreview bp;
            if (_cache.TryGetValue(key, out bp) && bp.Gray256 != null && bp.Binarized256 != null)
                return ClonePreview(bp);

            // Collect extents (geometry only; skip text/hatches/dims)
            Extents3d? extOpt = ComputeBlockExtentsFiltered(btr, tr);
            if (!extOpt.HasValue)
            {
                // Empty fallback: blank 256x256
                bp.Gray256 = new Bitmap(256, 256);
                using (var g = Graphics.FromImage(bp.Gray256)) g.Clear(Color.White);
                bp.Binarized256 = (Bitmap)bp.Gray256.Clone();
                _cache[key] = ClonePreview(bp);
                return bp;
            }

            var ext = extOpt.Value;

            // Padding
            double w = ext.MaxPoint.X - ext.MinPoint.X;
            double d = ext.MaxPoint.Y - ext.MinPoint.Y;
            double padW = w * PadFactor;
            double padD = d * PadFactor;

            var padded = new Extents3d(
                new Point3d(ext.MinPoint.X - padW, ext.MinPoint.Y - padD, 0),
                new Point3d(ext.MaxPoint.X + padW, ext.MaxPoint.Y + padD, 0));

            // Rasterize to 256x256, keep aspect, white bg
            Bitmap gray = new Bitmap(256, 256);
            using (var g = Graphics.FromImage(gray))
            {
                g.Clear(Color.White);
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                // Fit-to-canvas mapping
                double W = padded.MaxPoint.X - padded.MinPoint.X;
                double D = padded.MaxPoint.Y - padded.MinPoint.Y;
                if (W < 1e-9) W = 1.0;
                if (D < 1e-9) D = 1.0;
                double scale = Math.Min(232.0 / W, 232.0 / D); // 232 fits with margins
                double ox = (256.0 - W * scale) * 0.5;
                double oy = (256.0 - D * scale) * 0.5;

                // Mirror correction for odd parity: flip X so later hashing aligns
                if ((br.ScaleFactors.X < 0) ^ (br.ScaleFactors.Y < 0))
                {
                    g.TranslateTransform(128f, 0f);
                    g.ScaleTransform(-1f, 1f);
                    g.TranslateTransform(-128f, 0f);
                }

                // Draw exploded entities
                DBObjectCollection col = new DBObjectCollection();
                br.Explode(col);
                try
                {
                    int i;
                    for (i = 0; i < col.Count; i++)
                    {
                        Entity e = col[i] as Entity;
                        if (e == null) continue;
                        if (SkipEntity(e)) continue;
                        DrawEntity2D(g, e, padded, scale, ox, oy);
                    }
                }
                finally
                {
                    int i;
                    for (i = 0; i < col.Count; i++) col[i].Dispose();
                    col.Clear();
                }
            }

            // Binarize (Otsu) then thin
            Bitmap bin = BinarizeOtsu(gray);
            Bitmap thin = ThinningGuoHall(bin);

            bp.Gray256 = gray;
            bp.Binarized256 = thin;

            _cache[key] = ClonePreview(bp);
            return bp;
        }

        private static bool SkipEntity(Entity e)
        {
            // Filter out text/hatches/dimensions for preview pass
            if (e is DBText || e is MText) return true;
            if (e is Hatch) return true;
            if (e is Dimension) return true;
            if (e is Leader || e is MLeader) return true;
            return false;
        }

        private static Extents3d? ComputeBlockExtentsFiltered(BlockTableRecord btr, Transaction tr)
        {
            Extents3d? acc = null;
            foreach (ObjectId id in btr)
            {
                Entity e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (e == null || SkipEntity(e)) continue;
                try
                {
                    Extents3d ex = e.GeometricExtents;
                    if (!acc.HasValue) acc = ex;
                    else
                    {
                        Extents3d cur = acc.Value;
                        Point3d min = new Point3d(Math.Min(cur.MinPoint.X, ex.MinPoint.X),
                                                  Math.Min(cur.MinPoint.Y, ex.MinPoint.Y), 0);
                        Point3d max = new Point3d(Math.Max(cur.MaxPoint.X, ex.MaxPoint.X),
                                                  Math.Max(cur.MaxPoint.Y, ex.MaxPoint.Y), 0);
                        acc = new Extents3d(min, max);
                    }
                }
                catch { /* some entities may not have extents */ }
            }
            return acc;
        }

        private static void DrawEntity2D(Graphics g, Entity e, Extents3d padded, double s, double ox, double oy)
        {
            if (e is Line ln)
            {
                DrawLine(g, ln.StartPoint, ln.EndPoint, padded, s, ox, oy);
                return;
            }
            if (e is Polyline pl)
            {
                int n = pl.NumberOfVertices;
                int i;
                for (i = 0; i < n - 1; i++)
                {
                    Point2d p1 = pl.GetPoint2dAt(i);
                    Point2d p2 = pl.GetPoint2dAt(i + 1);
                    double bulge = pl.GetBulgeAt(i);
                    if (Math.Abs(bulge) > 1e-12)
                        DrawBulged(g, p1, p2, bulge, padded, s, ox, oy);
                    else
                        DrawLine(g, new Point3d(p1.X, p1.Y, 0), new Point3d(p2.X, p2.Y, 0), padded, s, ox, oy);
                }
                return;
            }
            if (e is Arc arc)
            {
                DrawArc(g, arc, padded, s, ox, oy);
                return;
            }
            if (e is Circle c)
            {
                DrawCircle(g, c, padded, s, ox, oy);
                return;
            }
            if (e is Ellipse el)
            {
                DrawEllipse(g, el, padded, s, ox, oy);
                return;
            }
            if (e is Spline sp)
            {
                int samples = 32;
                double t0 = sp.StartParam;
                double t1 = sp.EndParam;
                Point3d prev = sp.GetPointAtParameter(t0);
                int i;
                for (i = 1; i <= samples; i++)
                {
                    double t = t0 + (t1 - t0) * i / samples;
                    Point3d cur = sp.GetPointAtParameter(t);
                    DrawLine(g, prev, cur, padded, s, ox, oy);
                    prev = cur;
                }
                return;
            }
            if (e is BlockReference br)
            {
                DBObjectCollection parts = new DBObjectCollection();
                br.Explode(parts);
                try
                {
                    int i;
                    for (i = 0; i < parts.Count; i++)
                    {
                        Entity pe = parts[i] as Entity;
                        if (pe == null || SkipEntity(pe)) continue;
                        DrawEntity2D(g, pe, padded, s, ox, oy);
                    }
                }
                finally
                {
                    int i;
                    for (i = 0; i < parts.Count; i++) parts[i].Dispose();
                    parts.Clear();
                }
            }
        }

        private static void DrawLine(Graphics g, Point3d a, Point3d b, Extents3d pad, double s, double ox, double oy)
        {
            float x1 = (float)((a.X - pad.MinPoint.X) * s + ox);
            float y1 = (float)(256.0 - ((a.Y - pad.MinPoint.Y) * s + oy));
            float x2 = (float)((b.X - pad.MinPoint.X) * s + ox);
            float y2 = (float)(256.0 - ((b.Y - pad.MinPoint.Y) * s + oy));
            g.DrawLine(Pens.Black, x1, y1, x2, y2);
        }

        private static void DrawArc(Graphics g, Arc arc, Extents3d pad, double s, double ox, double oy)
        {
            double r = arc.Radius;
            float cx = (float)((arc.Center.X - pad.MinPoint.X) * s + ox);
            float cy = (float)(256.0 - ((arc.Center.Y - pad.MinPoint.Y) * s + oy));
            float rr = (float)(r * s);
            RectangleF rect = new RectangleF(cx - rr, cy - rr, rr * 2, rr * 2);
            float start = (float)(arc.StartAngle * 180.0 / Math.PI);
            float sweep = (float)((arc.EndAngle - arc.StartAngle) * 180.0 / Math.PI);
            g.DrawArc(Pens.Black, rect, start, sweep);
        }

        private static void DrawCircle(Graphics g, Circle c, Extents3d pad, double s, double ox, double oy)
        {
            float cx = (float)((c.Center.X - pad.MinPoint.X) * s + ox);
            float cy = (float)(256.0 - ((c.Center.Y - pad.MinPoint.Y) * s + oy));
            float rr = (float)(c.Radius * s);
            g.DrawEllipse(Pens.Black, cx - rr, cy - rr, rr * 2, rr * 2);
        }

        private static void DrawEllipse(Graphics g, Ellipse el, Extents3d pad, double s, double ox, double oy)
        {
            // Rough: draw bounding circle using major radius (preview only)
            float cx = (float)((el.Center.X - pad.MinPoint.X) * s + ox);
            float cy = (float)(256.0 - ((el.Center.Y - pad.MinPoint.Y) * s + oy));
            float rr = (float)(el.MajorRadius * s);
            g.DrawEllipse(Pens.Black, cx - rr, cy - rr, rr * 2, rr * 2);
        }

        private static void DrawBulged(Graphics g, Point2d p1, Point2d p2, double bulge, Extents3d pad, double s, double ox, double oy)
        {
            // Approximate with poly
            int segs = 10;
            int i;
            for (i = 0; i < segs; i++)
            {
                double t0 = (double)i / segs;
                double t1 = (double)(i + 1) / segs;
                Point2d q0 = InterpBulge(p1, p2, bulge, t0);
                Point2d q1 = InterpBulge(p1, p2, bulge, t1);
                DrawLine(g, new Point3d(q0.X, q0.Y, 0), new Point3d(q1.X, q1.Y, 0), pad, s, ox, oy);
            }
        }

        private static Point2d InterpBulge(Point2d p1, Point2d p2, double bulge, double t)
        {
            // Simple quadratic approximation on bulged arc
            double b = bulge;
            double x = (1 - t) * p1.X + t * p2.X + b * t * (1 - t) * (p2.Y - p1.Y);
            double y = (1 - t) * p1.Y + t * p2.Y + b * t * (1 - t) * (p1.X - p2.X);
            return new Point2d(x, y);
        }

        // --- Image ops ---

        private static Bitmap BinarizeOtsu(Bitmap src)
        {
            int w = src.Width, h = src.Height;
            int[] hist = new int[256];
            int x, y;

            for (y = 0; y < h; y++)
            {
                for (x = 0; x < w; x++)
                {
                    Color c = src.GetPixel(x, y);
                    int g = (int)(0.299 * c.R + 0.587 * c.G + 0.114 * c.B);
                    hist[g]++;
                }
            }

            int total = w * h;
            double sum = 0;
            for (x = 0; x < 256; x++) sum += x * hist[x];
            double sumB = 0;
            int wB = 0;
            int wF;
            double varMax = 0.0;
            int threshold = 127;

            for (x = 0; x < 256; x++)
            {
                wB += hist[x];
                if (wB == 0) continue;
                wF = total - wB;
                if (wF == 0) break;
                sumB += (double)(x * hist[x]);
                double mB = sumB / wB;
                double mF = (sum - sumB) / wF;
                double varBetween = (double)wB * (double)wF * (mB - mF) * (mB - mF);
                if (varBetween > varMax)
                {
                    varMax = varBetween;
                    threshold = x;
                }
            }

            Bitmap dst = new Bitmap(w, h);
            for (y = 0; y < h; y++)
            {
                for (x = 0; x < w; x++)
                {
                    Color c = src.GetPixel(x, y);
                    int g = (int)(0.299 * c.R + 0.587 * c.G + 0.114 * c.B);
                    dst.SetPixel(x, y, g > threshold ? Color.White : Color.Black);
                }
            }
            return dst;
        }

        // Guo-Hall thinning (simple, 256x256 => negligible cost)
        private static Bitmap ThinningGuoHall(Bitmap bin)
        {
            int w = bin.Width, h = bin.Height;
            byte[,] img = new byte[h, w];
            int x, y;

            for (y = 0; y < h; y++)
                for (x = 0; x < w; x++)
                    img[y, x] = (bin.GetPixel(x, y).R < 128) ? (byte)1 : (byte)0;

            bool changed;
            do
            {
                changed = false;
                changed |= ThinPass(img, w, h, 0);
                changed |= ThinPass(img, w, h, 1);
            } while (changed);

            Bitmap dst = new Bitmap(w, h);
            for (y = 0; y < h; y++)
                for (x = 0; x < w; x++)
                    dst.SetPixel(x, y, img[y, x] != 0 ? Color.Black : Color.White);
            return dst;
        }

        private static bool ThinPass(byte[,] p, int w, int h, int step)
        {
            bool changed = false;
            List<System.Drawing.Point> del = new List<System.Drawing.Point>(4096);
            int x, y;
            for (y = 1; y < h - 1; y++)
            {
                for (x = 1; x < w - 1; x++)
                {
                    if (p[y, x] == 0) continue;
                    byte p2 = p[y - 1, x]; byte p3 = p[y - 1, x + 1];
                    byte p4 = p[y, x + 1]; byte p5 = p[y + 1, x + 1];
                    byte p6 = p[y + 1, x]; byte p7 = p[y + 1, x - 1];
                    byte p8 = p[y, x - 1]; byte p9 = p[y - 1, x - 1];

                    int C = (~p2 & (p3 | p4)) + (~p4 & (p5 | p6)) + (~p6 & (p7 | p8)) + (~p8 & (p9 | p2));
                    int N1 = (p9 | p2) + (p3 | p4) + (p5 | p6) + (p7 | p8);
                    int N2 = (p2 | p3) + (p4 | p5) + (p6 | p7) + (p8 | p9);
                    int N = Math.Min(N1, N2);
                    int m = step == 0 ? ((p2 | p3 | ~p5) & p4) : ((p6 | p7 | ~p9) & p8);

                    if (C == 1 && (N >= 2 && N <= 3) && m == 0)
                        del.Add(new System.Drawing.Point(x, y));
                }
            }
            int i;
            for (i = 0; i < del.Count; i++)
            {
                System.Drawing.Point q = del[i];
                p[q.Y, q.X] = 0;
                changed = true;
            }
            return changed;
        }

        private static BlockPreview ClonePreview(BlockPreview bp)
        {
            return new BlockPreview
            {
                Gray256 = (Bitmap)bp.Gray256.Clone(),
                Binarized256 = (Bitmap)bp.Binarized256.Clone()
            };
        }
    }
}
