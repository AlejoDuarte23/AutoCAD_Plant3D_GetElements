using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.DataLinks;

[assembly: CommandClass(typeof(PlantJsonExporter.PlantJsonExportCommands))]

namespace PlantJsonExporter
{
    public sealed class PlantJsonExportCommands : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }

        [CommandMethod("EXPORTPLANTJSON", CommandFlags.Session)]
        public static void ExportPlantJson()
        {
            string projectXml = Environment.GetEnvironmentVariable("PLANT_PROJECT_XML") ?? "";
            string jsonOut = Environment.GetEnvironmentVariable("PLANT_JSON_OUT") ?? "";

            if (string.IsNullOrWhiteSpace(projectXml) || !File.Exists(projectXml))
                throw new InvalidOperationException("PLANT_PROJECT_XML is not set or does not exist.");

            if (string.IsNullOrWhiteSpace(jsonOut))
                throw new InvalidOperationException("PLANT_JSON_OUT is not set.");

            Directory.CreateDirectory(Path.GetDirectoryName(jsonOut)!);

            // Load Plant project from Project.xml (supported pattern used in Autodesk examples). 
            PlantProject plantPrj = PlantProject.LoadProject(projectXml, true, null, null);

            // Enumerate drawings from the project-managed list (not folder scan).
            List<string> drawingPaths = EnumerateProjectDrawings(plantPrj);

            var docMan = Application.DocumentManager;

            using var fs = new FileStream(jsonOut, FileMode.Create, FileAccess.Write, FileShare.Read);
            using var writer = new Utf8JsonWriter(fs, new JsonWriterOptions { Indented = true });

            writer.WriteStartObject();
            writer.WriteString("projectXml", projectXml);
            writer.WriteString("projectFolderPath", plantPrj.ProjectFolderPath);
            writer.WriteString("projectName", plantPrj.Name);
            writer.WriteString("exportedAtUtc", DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));

            writer.WriteStartArray("drawings");

            foreach (string dwgPath in drawingPaths)
            {
                WriteDrawingEntry(writer, docMan, plantPrj, dwgPath);
            }

            writer.WriteEndArray(); // drawings
            writer.WriteEndObject(); // root
            writer.Flush();

            plantPrj.Close();
        }

        private static List<string> EnumerateProjectDrawings(PlantProject plantPrj)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (object part in plantPrj.ProjectParts)
            {
                if (part is not Project prjPart) continue;

                try
                {
                    // GetPnPDrawingFiles + AbsoluteFileName are commonly used to list project drawings.
                    foreach (PnPProjectDrawing d in prjPart.GetPnPDrawingFiles())
                    {
                        if (!string.IsNullOrWhiteSpace(d.AbsoluteFileName))
                            set.Add(d.AbsoluteFileName);
                    }
                }
                catch
                {
                    // Some parts may not support listing or may throw; ignore and continue.
                }
            }

            return set.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static void WriteDrawingEntry(
            Utf8JsonWriter writer,
            DocumentCollection docMan,
            PlantProject plantPrj,
            string dwgPath)
        {
            writer.WriteStartObject();
            writer.WriteString("path", dwgPath);

            var (resolvedPath, resolutionNote) = ResolveDrawingPath(plantPrj.ProjectFolderPath, dwgPath);

            if (!string.Equals(resolvedPath, dwgPath, StringComparison.OrdinalIgnoreCase))
                writer.WriteString("resolvedPath", resolvedPath);

            if (!string.IsNullOrWhiteSpace(resolutionNote))
                writer.WriteString("pathResolution", resolutionNote);

            if (!File.Exists(resolvedPath))
            {
                writer.WriteString("error", "File not found on disk (not available locally).");
                writer.WriteEndObject();
                return;
            }

            Document? doc = null;

            try
            {
                doc = docMan.Open(resolvedPath, false);

                using (doc.LockDocument())
                {
                    ExportOpenedDrawing(writer, doc, plantPrj);
                }

                doc.CloseAndDiscard();
                writer.WriteEndObject();
            }
            catch (System.Exception ex)
            {
                try { doc?.CloseAndDiscard(); } catch { /* ignore */ }

                writer.WriteString("error", ex.Message);
                writer.WriteEndObject();
            }
        }

        private static (string resolvedPath, string? note) ResolveDrawingPath(string projectRoot, string pathFromProject)
        {
            if (File.Exists(pathFromProject))
                return (pathFromProject, null);

            if (string.IsNullOrWhiteSpace(projectRoot))
                return (pathFromProject, null);

            var root = Path.GetFullPath(projectRoot);

            // Split into parts using both separators
            var parts = pathFromProject
                .Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                    StringSplitOptions.RemoveEmptyEntries)
                .ToArray();

            // Try to anchor on a folder name that exists under the current project root
            // Example: ...\PID DWG\X.dwg -> <root>\PID DWG\X.dwg
            for (int i = parts.Length - 2; i >= 0; i--)
            {
                var dirName = parts[i];
                var candidateDir = Path.Combine(root, dirName);
                if (!Directory.Exists(candidateDir))
                    continue;

                var tail = Path.Combine(parts.Skip(i).ToArray());
                var candidate = Path.Combine(root, tail);

                if (File.Exists(candidate))
                    return (candidate, $"Remapped under project folder using '{dirName}'.");
            }

            // Fallback: search by filename inside the project folder
            var fileName = Path.GetFileName(pathFromProject);
            if (string.IsNullOrWhiteSpace(fileName))
                return (pathFromProject, null);

            try
            {
                var matches = Directory
                    .EnumerateFiles(root, fileName, SearchOption.AllDirectories)
                    .Take(10)
                    .ToList();

                if (matches.Count == 1)
                    return (matches[0], "Found by filename search under project folder.");

                if (matches.Count > 1)
                    return (pathFromProject, $"Multiple matches for '{fileName}' under project folder; not auto-resolving.");
            }
            catch
            {
                // ignore search errors
            }

            return (pathFromProject, null);
        }

        private static void ExportOpenedDrawing(Utf8JsonWriter writer, Document doc, PlantProject plantPrj)
        {
            var db = doc.Database;

            writer.WriteString("dwgVersion", db.OriginalFileVersion.ToString());
            writer.WriteString("insUnits", db.Insunits.ToString());

            // Prepare DataLinksManagers from all project parts that have them
            var dlms = GetAllDataLinksManagers(plantPrj);

            writer.WriteStartArray("entities");

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in ms)
                {
                    if (!id.IsValid) continue;

                    Entity? ent;
                    try { ent = tr.GetObject(id, OpenMode.ForRead) as Entity; }
                    catch { continue; }

                    if (ent == null) continue;

                    writer.WriteStartObject();

                    WriteEntityBasics(writer, ent);

                    if (ent is BlockReference br)
                    {
                        WriteBlockDetails(writer, tr, br);
                    }

                    WriteXData(writer, ent);

                    // Plant properties via DataLinks: FindAcPpRowId + GetAllProperties(rowId, true).
                    WritePlantLinks(writer, ent.ObjectId, dlms);

                    writer.WriteEndObject();
                }

                tr.Commit();
            }

            writer.WriteEndArray(); // entities
        }

        private static List<(string partName, DataLinksManager dlm)> GetAllDataLinksManagers(PlantProject plantPrj)
        {
            var list = new List<(string partName, DataLinksManager dlm)>();

            foreach (object part in plantPrj.ProjectParts)
            {
                if (part is not Project prjPart) continue;

                try
                {
                    if (prjPart.DataLinksManager != null)
                        list.Add((GetProjectPartName(prjPart), prjPart.DataLinksManager));
                }
                catch { }
            }

            return list;
        }

        private static void WriteEntityBasics(Utf8JsonWriter writer, Entity ent)
        {
            writer.WriteString("handle", ent.Handle.ToString());
            writer.WriteString("dxfName", ent.GetRXClass().DxfName);
            writer.WriteString("layer", ent.Layer);
            writer.WriteString("linetype", ent.Linetype);
            writer.WriteNumber("lineWeight", (int)ent.LineWeight);

            try
            {
                var c = ent.Color;
                writer.WriteStartObject("trueColor");
                writer.WriteNumber("r", c.Red);
                writer.WriteNumber("g", c.Green);
                writer.WriteNumber("b", c.Blue);
                writer.WriteEndObject();
            }
            catch { }

            try
            {
                writer.WriteNumber("transparency", ent.Transparency.Alpha);
            }
            catch { }

            try
            {
                var ext = ent.GeometricExtents;
                writer.WriteStartObject("extents");
                writer.WriteStartArray("min");
                writer.WriteNumberValue(ext.MinPoint.X);
                writer.WriteNumberValue(ext.MinPoint.Y);
                writer.WriteNumberValue(ext.MinPoint.Z);
                writer.WriteEndArray();
                writer.WriteStartArray("max");
                writer.WriteNumberValue(ext.MaxPoint.X);
                writer.WriteNumberValue(ext.MaxPoint.Y);
                writer.WriteNumberValue(ext.MaxPoint.Z);
                writer.WriteEndArray();
                writer.WriteEndObject();
            }
            catch { }
        }

        private static void WriteBlockDetails(Utf8JsonWriter writer, Transaction tr, BlockReference br)
        {
            writer.WriteStartObject("block");

            try
            {
                var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                writer.WriteString("name", btr.Name);
            }
            catch { }

            try
            {
                if (br.IsDynamicBlock)
                {
                    var dynBtr = (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                    writer.WriteString("dynamicBaseName", dynBtr.Name);
                }
            }
            catch { }

            writer.WriteStartArray("attributes");
            foreach (ObjectId attId in br.AttributeCollection)
            {
                if (!attId.IsValid) continue;
                if (tr.GetObject(attId, OpenMode.ForRead) is AttributeReference ar)
                {
                    writer.WriteStartObject();
                    writer.WriteString("tag", ar.Tag);
                    writer.WriteString("text", ar.TextString);
                    writer.WriteEndObject();
                }
            }
            writer.WriteEndArray();

            // Dynamic block properties
            try
            {
                var props = br.DynamicBlockReferencePropertyCollection;
                writer.WriteStartArray("dynamicProperties");

                foreach (dynamic p in props)
                {
                    try
                    {
                        writer.WriteStartObject();
                        writer.WriteString("name", (string)p.PropertyName);
                        writer.WriteString("value", p.Value?.ToString() ?? "");
                        writer.WriteBoolean("readOnly", (bool)p.ReadOnly);
                        writer.WriteEndObject();
                    }
                    catch { }
                }

                writer.WriteEndArray();
            }
            catch { }

            writer.WriteEndObject(); // block
        }

        private static void WriteXData(Utf8JsonWriter writer, Entity ent)
        {
            try
            {
                ResultBuffer? rb = ent.XData;
                if (rb == null) return;

                writer.WriteStartArray("xdata");
                foreach (TypedValue tv in rb)
                {
                    writer.WriteStartObject();
                    writer.WriteNumber("typeCode", tv.TypeCode);
                    writer.WriteString("value", tv.Value?.ToString() ?? "");
                    writer.WriteEndObject();
                }
                writer.WriteEndArray();
            }
            catch { }
        }

        private static void WritePlantLinks(
            Utf8JsonWriter writer,
            ObjectId objId,
            List<(string partName, DataLinksManager dlm)> dlms)
        {
            writer.WriteStartArray("plantLinks");

            foreach (var (partName, dlm) in dlms)
            {
                try
                {
                    int rowId = dlm.FindAcPpRowId(objId); //
                    if (rowId <= 0) continue;

                    var props = dlm.GetAllProperties(rowId, true);

                    writer.WriteStartObject();
                    writer.WriteString("projectPart", partName);
                    writer.WriteNumber("rowId", rowId);

                    writer.WriteStartObject("properties");
                    WriteProperties(writer, props);
                    writer.WriteEndObject();

                    writer.WriteEndObject();
                }
                catch
                {
                    // Not linked in this DLM (normal for many objects)
                }
            }

            writer.WriteEndArray();
        }

        private static string GetProjectPartName(Project prjPart)
        {
            var type = prjPart.GetType();
            var prop = type.GetProperty("Name") ?? type.GetProperty("PartName") ?? type.GetProperty("ProjectPartName");
            if (prop != null && prop.PropertyType == typeof(string))
            {
                return (string?)prop.GetValue(prjPart) ?? type.Name;
            }

            return type.Name;
        }

        private static void WriteProperties(Utf8JsonWriter writer, object? props)
        {
            if (props is NameValueCollection nvc)
            {
                foreach (string key in nvc.AllKeys)
                {
                    if (!string.IsNullOrEmpty(key))
                        writer.WriteString(key, nvc[key] ?? "");
                }
                return;
            }

            if (props is IEnumerable<KeyValuePair<string, string>> kvps)
            {
                foreach (var kvp in kvps)
                {
                    if (!string.IsNullOrEmpty(kvp.Key))
                        writer.WriteString(kvp.Key, kvp.Value ?? "");
                }
                return;
            }

            if (props is IDictionary dict)
            {
                foreach (DictionaryEntry entry in dict)
                {
                    var key = entry.Key as string;
                    if (!string.IsNullOrEmpty(key))
                        writer.WriteString(key, entry.Value?.ToString() ?? "");
                }
            }
        }
    }
}
