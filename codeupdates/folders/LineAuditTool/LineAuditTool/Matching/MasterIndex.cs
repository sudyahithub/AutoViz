// MasterIndex.cs
// Minimal cache for masters with precomputed features.
// Uses Newtonsoft.Json for simplicity (available in most AutoCAD plugin setups).
// C# 7.3.

using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Autodesk.AutoCAD.Geometry;


namespace LineAuditTool.Matching
{
    public sealed class MasterIndex
    {
        public List<MasterEntry> Masters = new List<MasterEntry>();

        public static MasterIndex LoadOrBuild(Database masterDb, string cachePath)
        {
            try
            {
                if (!string.IsNullOrEmpty(cachePath) && File.Exists(cachePath))
                {
                    string json = File.ReadAllText(cachePath);
                    var idx = JsonConvert.DeserializeObject<MasterIndex>(json);
                    if (idx != null && idx.Masters != null && idx.Masters.Count > 0)
                        return idx;
                }
            }
            catch (System.Exception) { /* ignore */ }

            // Build minimal index by scanning masterDb block defs
            MasterIndex mi = new MasterIndex();
            using (Transaction tr = masterDb.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(masterDb.BlockTableId, OpenMode.ForRead);
                PreviewRenderer prv = new PreviewRenderer();

                foreach (ObjectId id in bt)
                {
                    BlockTableRecord btr = tr.GetObject(id, OpenMode.ForRead) as BlockTableRecord;
                    if (btr == null || btr.IsAnonymous || btr.IsLayout) continue;

                    // Preview + hashes
                    // Fake BlockReference to carry scale parity = +1 (no mirror)
                    BlockReference dummy = new BlockReference(Point3d.Origin, id);
                    BlockPreview bp = prv.GetPreview(dummy, tr);
                    HashFeatures hf = ImageHash.ComputeHashes(bp.Binarized256);

                    // Geometry
                    GeomFeatures gf = GeometrySignature.Compute(btr, tr);

                    MasterEntry me = new MasterEntry();
                    me.Name = btr.Name;
                    me.Category = ""; // plug your taxonomy if available
                    me.WidthMM = gf.WidthMM;
                    me.DepthMM = gf.DepthMM;
                    me.Hash = hf;
                    me.Geom = gf;

                    mi.Masters.Add(me);

                    if (bp.Gray256 != null) bp.Gray256.Dispose();
                    if (bp.Binarized256 != null) bp.Binarized256.Dispose();
                }
                tr.Commit();
            }

            try
            {
                if (!string.IsNullOrEmpty(cachePath))
                {
                    string dir = Path.GetDirectoryName(cachePath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    File.WriteAllText(cachePath, JsonConvert.SerializeObject(mi, Formatting.Indented));
                }
            }
            catch (System.Exception) { /* ignore */ }

            return mi;
        }
    }
}
