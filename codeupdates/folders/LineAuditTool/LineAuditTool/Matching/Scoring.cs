// Scoring.cs
// Combined scoring and RankCandidates API.
// C# 7.3.

using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;

namespace LineAuditTool.Matching
{
    public static class Scoring
    {
        // Rank candidates from a list of master entries (with precomputed features).
        // Optional attribute prefilter: category + width ±20% if available.
        public static List<CandidateScore> RankCandidates(
            BlockReference source,
            Transaction tr,
            IEnumerable<MasterEntry> masters,
            MatchConfig cfg,
            PreviewRenderer previewRenderer // reuse instance to leverage its cache
        )
        {
            if (cfg == null) cfg = new MatchConfig();
            if (previewRenderer == null) previewRenderer = new PreviewRenderer();

            // --- Source preview & features ---
            BlockPreview sp;
            try { sp = previewRenderer.GetPreview(source, tr); }
            catch (System.Exception) { sp = new BlockPreview(); sp.Gray256 = new System.Drawing.Bitmap(256, 256); sp.Binarized256 = (System.Drawing.Bitmap)sp.Gray256.Clone(); }

            HashFeatures sHash = ImageHash.ComputeHashes(sp.Binarized256);

            // Source geometry from definition (parity doesn't matter here)
            BlockTableRecord srcBtr = (BlockTableRecord)tr.GetObject(source.BlockTableRecord, OpenMode.ForRead);
            GeomFeatures sGeom = GeometrySignature.Compute(srcBtr, tr);

            // Prefilter settings (optional)
            string srcCategory = null;      // TODO: plug your own source category if known
            double srcWidth = sGeom.WidthMM;
            double wMin = srcWidth * 0.8, wMax = srcWidth * 1.2;

            // --- Score loop ---
            List<CandidateScore> scores = new List<CandidateScore>(128);
            foreach (MasterEntry m in masters)
            {
                if (m == null) continue;

                // Prefilter (optional)
                if (!string.IsNullOrEmpty(srcCategory) && !string.IsNullOrEmpty(m.Category))
                {
                    if (!string.Equals(srcCategory, m.Category, StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                if (m.WidthMM > 0.0 && (m.WidthMM < wMin || m.WidthMM > wMax))
                {
                    // pass if width known and grossly out of range
                    continue;
                }

                // Hash distances
                int dD = ImageHash.Hamming64(sHash.DHash64, m.Hash.DHash64);
                int dP = ImageHash.Hamming64(sHash.PHash64, m.Hash.PHash64);

                double dNorm = (double)dD / 64.0;
                double pNorm = (double)dP / 64.0;

                double hashSim = cfg.HashBlend.P * (1.0 - pNorm) + cfg.HashBlend.D * (1.0 - dNorm);
                if (hashSim < 0.0) hashSim = 0.0; if (hashSim > 1.0) hashSim = 1.0;

                // Geom similarity
                GeomFeatures g = m.Geom;
                double geomSim = GeometrySignature.GeomSimilarity(ref sGeom, ref g);

                double finalScore = cfg.Weights.Hash * hashSim + cfg.Weights.Geometry * geomSim;
                if (finalScore < 0.0) finalScore = 0.0; if (finalScore > 1.0) finalScore = 1.0;

                CandidateScore cs;
                cs.MasterName = m.Name;
                cs.Score = finalScore;
                cs.HashSim = hashSim;
                cs.GeomSim = geomSim;
                cs.PHashHamming = dP;
                cs.DHashHamming = dD;
                scores.Add(cs);
            }

            // Partial selection sort to MaxCandidates
            int maxN = cfg.MaxCandidates > 0 ? cfg.MaxCandidates : 50;
            if (scores.Count > 1)
            {
                int i, j, best;
                int limit = Math.Min(scores.Count, maxN);
                for (i = 0; i < limit; i++)
                {
                    best = i;
                    for (j = i + 1; j < scores.Count; j++)
                        if (scores[j].Score > scores[best].Score) best = j;
                    if (best != i)
                    {
                        CandidateScore tmp = scores[i];
                        scores[i] = scores[best];
                        scores[best] = tmp;
                    }
                }
                if (scores.Count > limit) scores.RemoveRange(limit, scores.Count - limit);
            }

            // Filter by MinScore
            double minScore = cfg.MinScore;
            int k = 0;
            while (k < scores.Count)
            {
                if (scores[k].Score < minScore) scores.RemoveAt(k);
                else k++;
            }

            // Dispose previews (caller may reuse renderer cache for masters)
            if (sp.Gray256 != null) sp.Gray256.Dispose();
            if (sp.Binarized256 != null) sp.Binarized256.Dispose();

            return scores;
        }
    }
}
