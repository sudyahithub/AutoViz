// BlockClickMatcher.cs
// Build x64, target the .NET Framework your AutoCAD uses (e.g., 4.7.x).
// NuGet: Newtonsoft.Json

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Newtonsoft.Json;

// Avoid 'Application' ambiguity
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(LineAuditTool.BlockClickMatcher))]

namespace LineAuditTool
{
    public class BlockClickMatcher
    {
        // ========= CONFIG: PATHS =========
        private const string MatchesByHandleJsonPath =
            @"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS\_MATCH_RESULTS\work_to_master.json";

        private const string MatchesByStemJsonPath =
            @"C:\Users\admin\Downloads\VIZ-AUTOCAD\EXPORTS\_MATCH_RESULTS\work_to_master_by_stem.json";

        private const string MasterDwgPath =
            @"C:\Users\admin\Downloads\VIZ-AUTOCAD\M1.dwg";
        // =================================

        // ===== JSON DTOs =====
        private class TopK
        {
            [JsonProperty("master_block")] public string MasterBlock { get; set; }
            [JsonProperty("score")] public double Score { get; set; }
            [JsonProperty("preview")] public string Preview { get; set; }
            [JsonProperty("rot")] public int RotationDeltaDeg { get; set; } // 0/90/180/270
        }
        private class MatchEntry
        {
            [JsonProperty("work_image")] public string WorkImage { get; set; }
            [JsonProperty("topk")] public List<TopK> TopK { get; set; }
        }

        // caches
        private static Dictionary<string, MatchEntry> _byHandle;
        private static Dictionary<string, MatchEntry> _byStem;
        private static DateTime _byHandleMtimeUtc;
        private static DateTime _byStemMtimeUtc;

        private static Dictionary<string, MatchEntry> LoadMap(string path, ref Dictionary<string, MatchEntry> cache, ref DateTime ts)
        {
            if (!File.Exists(path))
                return new Dictionary<string, MatchEntry>(StringComparer.OrdinalIgnoreCase);

            var mtime = File.GetLastWriteTimeUtc(path);
            if (cache != null && mtime == ts)
                return cache;

            var json = File.ReadAllText(path);
            var raw = JsonConvert.DeserializeObject<Dictionary<string, MatchEntry>>(json)
                      ?? new Dictionary<string, MatchEntry>(StringComparer.OrdinalIgnoreCase);

            var norm = new Dictionary<string, MatchEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in raw)
            {
                var k = kv.Key?.Trim();
                if (string.IsNullOrEmpty(k)) continue;
                if (path.EndsWith("work_to_master.json", StringComparison.OrdinalIgnoreCase))
                {
                    if (k.StartsWith("H", StringComparison.OrdinalIgnoreCase)) k = k.Substring(1);
                    k = k.ToUpperInvariant();
                }
                norm[k] = kv.Value;
            }
            cache = norm;
            ts = mtime;
            return cache;
        }

        private static Dictionary<string, MatchEntry> ByHandle() => LoadMap(MatchesByHandleJsonPath, ref _byHandle, ref _byHandleMtimeUtc);
        private static Dictionary<string, MatchEntry> ByStem() => LoadMap(MatchesByStemJsonPath, ref _byStem, ref _byStemMtimeUtc);

        // ========= Helpers (naming, indices) =========

        private static string SafeName(string s)
        {
            string t = Regex.Replace(s ?? "", @"[<>:""/\\|?*]+", "_").Trim();
            return string.IsNullOrEmpty(t) ? "Unnamed" : t;
        }
        private static string StripIndexSuffix(string s) => Regex.Replace(s ?? "", @"__\d{2,}$", "");
        private static string NormalizeForCompare(string s) =>
            SafeName(StripIndexSuffix(s)).Replace(' ', '_').ToLowerInvariant();

        private static long HandleToLong(Handle h)
        {
            return long.TryParse(h.ToString(), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var v)
                ? v : long.MaxValue;
        }

        // Build "<BlockName>__NNN" where NNN is the index of this insert among peers, sorted by HANDLE
        private static string BuildStemKey(Document doc, Transaction tr, BlockReference br)
        {
            var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
            string blockName = btr.Name;

            string s = SafeName(blockName);

            var owner = (BlockTableRecord)tr.GetObject(br.OwnerId, OpenMode.ForRead);
            var peers = new List<BlockReference>();
            foreach (ObjectId id in owner)
            {
                if (id.ObjectClass.DxfName != "INSERT") continue;
                var br2 = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                if (br2 == null) continue;
                var btr2 = (BlockTableRecord)tr.GetObject(br2.BlockTableRecord, OpenMode.ForRead);
                if (!string.Equals(btr2.Name, blockName, StringComparison.OrdinalIgnoreCase)) continue;
                peers.Add(br2);
            }
            peers.Sort((a, b) => HandleToLong(a.Handle).CompareTo(HandleToLong(b.Handle)));

            int idx = peers.FindIndex(p => p.ObjectId == br.ObjectId);
            if (idx < 0) idx = 0;

            return $"{s}__{(idx + 1):000}";
        }

        private static double Deg2Rad(double d) => d * Math.PI / 180.0;

        private static bool TryCenter(Entity e, out Point3d center)
        {
            center = Point3d.Origin;
            try
            {
                var ext = e.GeometricExtents;
                center = new Point3d(
                    (ext.MinPoint.X + ext.MaxPoint.X) * 0.5,
                    (ext.MinPoint.Y + ext.MaxPoint.Y) * 0.5,
                    (ext.MinPoint.Z + ext.MaxPoint.Z) * 0.5
                );
                return true;
            }
            catch { return false; }
        }

        // ========= Public Commands =========

        [CommandMethod("CLICKMATCH")]
        public static void Command_ClickMatch()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect a block to get MASTER recommendation:");
            peo.SetRejectMessage("\nOnly BlockReference is supported.");
            peo.AddAllowedClass(typeof(BlockReference), true);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var br = tr.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;
                if (br == null) { ed.WriteMessage("\nNot a BlockReference."); return; }

                // 1) Try by HANDLE
                var handleKey = br.Handle.ToString().ToUpperInvariant();
                if (handleKey.StartsWith("H")) handleKey = handleKey.Substring(1);
                var mapH = ByHandle();
                MatchEntry entry = null;
                if (!mapH.TryGetValue(handleKey, out entry))
                {
                    // 2) Fallback: by STEM "<BlockName>__NNN"
                    var stemKey = BuildStemKey(doc, tr, br);
                    var mapS = ByStem();
                    if (!mapS.TryGetValue(stemKey, out entry))
                    {
                        ed.WriteMessage($"\nNo match candidates for handle {handleKey} or stem '{stemKey}'.");
                        return;
                    }
                }

                var choice = ShowRecommendationsUI(entry);
                if (choice == null) return;

                ReplaceWithMasterBlock(doc, tr, br, choice.MasterBlock, choice.RotationDeltaDeg);
                tr.Commit();
            }
        }

        // Optional: debug helper to verify keys
        [CommandMethod("SHOWSTEM")]
        public static void Command_ShowStem()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            var peo = new PromptEntityOptions("\nPick a block to print its keys:");
            peo.SetRejectMessage("\nOnly BlockReference is supported.");
            peo.AddAllowedClass(typeof(BlockReference), true);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (doc.LockDocument())
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var br = tr.GetObject(per.ObjectId, OpenMode.ForRead) as BlockReference;
                if (br == null) { ed.WriteMessage("\nNot a BlockReference."); return; }

                var stem = BuildStemKey(doc, tr, br);
                var h = br.Handle.ToString().ToUpperInvariant();
                if (h.StartsWith("H")) h = h.Substring(1);

                ed.WriteMessage($"\nStem key: {stem}\nHandle key: {h}");
            }
        }

        // ========= UI =========

        private class Choice
        {
            public string MasterBlock;
            public double Score;
            public string PreviewPath;
            public int RotationDeltaDeg;
        }

        private static Choice ShowRecommendationsUI(MatchEntry entry)
        {
            var topk = entry?.TopK ?? new List<TopK>();
            if (topk.Count == 0) return null;

            var choices = topk.Select(k => new Choice
            {
                MasterBlock = k.MasterBlock,
                Score = k.Score,
                PreviewPath = k.Preview,
                RotationDeltaDeg = k.RotationDeltaDeg
            }).ToList();

            using (var dlg = new Form())
            {
                dlg.Text = "Master Block Recommendation";
                dlg.StartPosition = FormStartPosition.CenterScreen;
                dlg.Width = 760;
                dlg.Height = 560;

                var left = new ListBox { Dock = DockStyle.Left, Width = 320 };
                var rightPanel = new Panel { Dock = DockStyle.Fill };
                var preview = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = System.Drawing.Color.Black };
                var lblScore = new Label { Dock = DockStyle.Top, Height = 22, TextAlign = System.Drawing.ContentAlignment.MiddleCenter };
                var bottom = new Panel { Dock = DockStyle.Bottom, Height = 64 };
                var btnOk = new Button { Text = "Replace with selected", Dock = DockStyle.Right, Width = 220 };
                var btnCancel = new Button { Text = "Cancel", Dock = DockStyle.Right, Width = 120 };

                bottom.Controls.Add(btnCancel);
                bottom.Controls.Add(btnOk);
                rightPanel.Controls.Add(preview);
                rightPanel.Controls.Add(lblScore);

                dlg.Controls.Add(rightPanel);
                dlg.Controls.Add(left);
                dlg.Controls.Add(bottom);

                foreach (var c in choices)
                    left.Items.Add($"{c.MasterBlock}   (score {c.Score:0.000})");

                left.SelectedIndexChanged += (s, e) =>
                {
                    var idx = left.SelectedIndex;
                    if (idx < 0) return;
                    var c = choices[idx];
                    lblScore.Text = $"Match score: {c.Score:0.000}";
                    if (!string.IsNullOrWhiteSpace(c.PreviewPath) && File.Exists(c.PreviewPath))
                    {
                        try { preview.Image = System.Drawing.Image.FromFile(c.PreviewPath); }
                        catch { preview.Image = null; }
                    }
                    else preview.Image = null;
                };

                if (left.Items.Count > 0) left.SelectedIndex = 0;

                Choice result = null;
                btnOk.Click += (s, e) =>
                {
                    var i = left.SelectedIndex;
                    if (i >= 0)
                    {
                        result = choices[i];
                        dlg.DialogResult = DialogResult.OK;
                        dlg.Close();
                    }
                };
                btnCancel.Click += (s, e) => { dlg.DialogResult = DialogResult.Cancel; dlg.Close(); };

                var dr = AcApp.ShowModalDialog(dlg);
                return (dr == DialogResult.OK) ? result : null;
            }
        }

        // ========= Replacement (rotation delta + center-align pivot) =========

        private static void ReplaceWithMasterBlock(Document doc, Transaction tr, BlockReference oldBr, string requestedMasterName, int rotationDeltaDeg)
        {
            var db = doc.Database;
            var ed = doc.Editor;

            if (string.IsNullOrWhiteSpace(requestedMasterName))
            {
                ed.WriteMessage("\nInvalid master block name.");
                return;
            }

            var targetBtrId = EnsureMasterBlockDefinition(db, tr, requestedMasterName);
            if (targetBtrId == ObjectId.Null)
            {
                ed.WriteMessage($"\nCannot find or import master block '{requestedMasterName}'.");
                return;
            }

            var pos = oldBr.Position;
            var rot = oldBr.Rotation + Deg2Rad(rotationDeltaDeg);
            var scale = oldBr.ScaleFactors;
            var layerId = oldBr.LayerId;
            var normal = oldBr.Normal;

            var oldCenterOk = TryCenter(oldBr, out var oldCenter);
            var oldAttrs = CaptureAttributes(tr, oldBr);

            var newBr = new BlockReference(pos, targetBtrId)
            {
                Rotation = rot,
                Normal = normal
            };
            newBr.SetDatabaseDefaults();
            newBr.ScaleFactors = scale;
            newBr.LayerId = layerId;

            var curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            curSpace.AppendEntity(newBr);
            tr.AddNewlyCreatedDBObject(newBr, true);

            TrySyncAttributes(tr, newBr, oldAttrs);

            // Center-align to counter different base points
            if (oldCenterOk && TryCenter(newBr, out var newCenter))
            {
                var delta = oldCenter - newCenter;
                if (!delta.IsZeroLength())
                    newBr.TransformBy(Matrix3d.Displacement(delta));
            }

            oldBr.UpgradeOpen();
            oldBr.Erase();

            ed.WriteMessage($"\nReplaced with master '{requestedMasterName}' (rotΔ={rotationDeltaDeg}°).");
        }

        // === SMART IMPORT: map preview names to real block names in M1.dwg ===
        private static ObjectId EnsureMasterBlockDefinition(Database targetDb, Transaction tr, string requestedName)
        {
            var bt = (BlockTable)tr.GetObject(targetDb.BlockTableId, OpenMode.ForRead);

            string[] localCandidates =
            {
                requestedName,
                StripIndexSuffix(requestedName),
                SafeName(requestedName),
                SafeName(StripIndexSuffix(requestedName))
            };
            foreach (var cand in localCandidates)
            {
                if (!string.IsNullOrWhiteSpace(cand) && bt.Has(cand))
                    return bt[cand];
            }

            if (!File.Exists(MasterDwgPath)) return ObjectId.Null;

            string importedName = null;

            using (var masterDb = new Database(false, true))
            {
                masterDb.ReadDwgFile(MasterDwgPath, FileShare.Read, allowCPConversion: true, password: "");
                using (var trM = masterDb.TransactionManager.StartTransaction())
                {
                    var btM = (BlockTable)trM.GetObject(masterDb.BlockTableId, OpenMode.ForRead);

                    foreach (var cand in localCandidates)
                    {
                        if (!string.IsNullOrWhiteSpace(cand) && btM.Has(cand))
                        {
                            importedName = cand;
                            var ids = new ObjectIdCollection(new[] { btM[cand] });
                            var idMap = new IdMapping();
                            masterDb.WblockCloneObjects(ids, targetDb.BlockTableId, idMap,
                                DuplicateRecordCloning.Replace, deferTranslation: false);
                            trM.Commit();
                            goto IMPORT_DONE;
                        }
                    }

                    string targetNorm = NormalizeForCompare(requestedName);
                    string fallbackName = null;

                    foreach (ObjectId id in btM)
                    {
                        var btrM = (BlockTableRecord)trM.GetObject(id, OpenMode.ForRead);
                        if (btrM.IsLayout) continue;
                        string realName = btrM.Name;
                        string norm = NormalizeForCompare(realName);

                        if (string.Equals(norm, targetNorm, StringComparison.OrdinalIgnoreCase))
                        {
                            importedName = realName;
                            var ids = new ObjectIdCollection(new[] { id });
                            var idMap = new IdMapping();
                            masterDb.WblockCloneObjects(ids, targetDb.BlockTableId, idMap,
                                DuplicateRecordCloning.Replace, deferTranslation: false);
                            trM.Commit();
                            goto IMPORT_DONE;
                        }
                        if (fallbackName == null && norm.StartsWith(targetNorm, StringComparison.OrdinalIgnoreCase))
                            fallbackName = realName;
                    }

                    if (fallbackName != null)
                    {
                        importedName = fallbackName;
                        var ids = new ObjectIdCollection(new[] { btM[importedName] });
                        var idMap = new IdMapping();
                        masterDb.WblockCloneObjects(ids, targetDb.BlockTableId, idMap,
                            DuplicateRecordCloning.Replace, deferTranslation: false);
                        trM.Commit();
                    }
                }
            }

IMPORT_DONE:
            bt = (BlockTable)tr.GetObject(targetDb.BlockTableId, OpenMode.ForRead);

            foreach (var cand in localCandidates)
            {
                if (!string.IsNullOrWhiteSpace(cand) && bt.Has(cand))
                    return bt[cand];
            }
            if (!string.IsNullOrWhiteSpace(importedName) && bt.Has(importedName))
                return bt[importedName];

            string targetNorm2 = NormalizeForCompare(requestedName);
            foreach (ObjectId id in bt)
            {
                var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                if (btr.IsLayout) continue;
                if (string.Equals(NormalizeForCompare(btr.Name), targetNorm2, StringComparison.OrdinalIgnoreCase))
                    return id;
            }
            return ObjectId.Null;
        }

        // ========= Attributes helpers =========

        private static Dictionary<string, string> CaptureAttributes(Transaction tr, BlockReference br)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (br.AttributeCollection == null) return result;
                foreach (ObjectId aid in br.AttributeCollection)
                {
                    if (!aid.IsValid) continue;
                    var obj = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                    if (obj == null) continue;
                    var tag = obj.Tag;
                    var val = obj.TextString;
                    if (!string.IsNullOrEmpty(tag))
                        result[tag] = val ?? string.Empty;
                }
            }
            catch { /* ignore */ }
            return result;
        }

        private static void TrySyncAttributes(Transaction tr, BlockReference newBr, Dictionary<string, string> oldAttrs)
        {
            if (oldAttrs == null) oldAttrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var btr = (BlockTableRecord)tr.GetObject(newBr.BlockTableRecord, OpenMode.ForRead);
            if (!btr.HasAttributeDefinitions) return;

            foreach (ObjectId id in btr)
            {
                if (!(tr.GetObject(id, OpenMode.ForRead) is AttributeDefinition ad)) continue;
                if (ad.Constant) continue;

                var ar = new AttributeReference();
                ar.SetAttributeFromBlock(ad, newBr.BlockTransform);

                if (oldAttrs.TryGetValue(ad.Tag, out var v) && !string.IsNullOrEmpty(v))
                    ar.TextString = v;
                else
                    ar.TextString = ad.TextString;

                newBr.AttributeCollection.AppendAttribute(ar);
                tr.AddNewlyCreatedDBObject(ar, true);
            }
        }
    }
}
