// Models.cs
// DTOs and tiny helpers used across the matching pipeline.
// C# 7.3 compatible.

using System;
using System.Drawing;

namespace LineAuditTool.Matching
{
    public struct BlockPreview
    {
        public Bitmap Gray256;      // 256x256 8-bit grayscale (stored as 24bpp RGB for GDI+)
        public Bitmap Binarized256; // 256x256 binarized (0/255)
        // NOTE: caller owns disposal.
    }

    public struct HashFeatures
    {
        public ulong DHash64;
        public ulong PHash64;
    }

    public struct GeomFeatures
    {
        public double WidthMM;
        public double DepthMM;
        public double Aspect;   // Width / Depth (safe eps handled by producer)

        public int LineCount;
        public int ArcCount;
        public int SplineCount;
        public int PolyCount;

        public int Components;      // approximate via bbox proximity union-find
        public int[] AngleHist8;    // length 8, sum normalized to 1.0 (ints scaled by 1000 OK)
    }

    public struct CandidateScore
    {
        public string MasterName;
        public double Score;          // 0..1
        public double HashSim;        // 0..1
        public double GeomSim;        // 0..1
        public int PHashHamming;      // 0..64
        public int DHashHamming;      // 0..64
    }

    public sealed class MasterEntry
    {
        public string Name;
        public string Category;             // optional, empty/null ok
        public double WidthMM;
        public double DepthMM;

        public HashFeatures Hash;
        public GeomFeatures Geom;
    }

    public sealed class MatchConfig
    {
        public WeightsConfig Weights = new WeightsConfig();
        public HashBlendConfig HashBlend = new HashBlendConfig();

        public int MaxCandidates = 50;
        public double AmbiguityMargin = 0.08; // not used here, reserved for UI tie-breaks
        public double MinScore = 0.65;

        public sealed class WeightsConfig
        {
            public double Hash = 0.7;
            public double Geometry = 0.3;
        }

        public sealed class HashBlendConfig
        {
            public double P = 0.6;
            public double D = 0.4;
        }
    }

    internal static class Util
    {
        public const double EPS = 1e-9;

        public static double Clamp01(double v)
        {
            if (v < 0.0) return 0.0;
            if (v > 1.0) return 1.0;
            return v;
        }

        public static double SafeDiv(double a, double b)
        {
            if (Math.Abs(b) < EPS) return 0.0;
            return a / b;
        }
    }
}
