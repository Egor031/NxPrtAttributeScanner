using NXOpen;
using System;
using System.Collections.Generic;
using System.IO;

public static class Scanner
{
    public static void Run(Session s, ScanOptions opt, CacheRepository repo)
    {
        if (string.IsNullOrWhiteSpace(opt.RootFolder) || !Directory.Exists(opt.RootFolder))
            throw new Exception("Root folder not found: " + (opt.RootFolder ?? "<null>"));

        if (opt.Mode == RunMode.ScanAndExport)
        {
            if (string.IsNullOrWhiteSpace(opt.ExcelOutputPath))
                throw new Exception("ExcelOutputPath is empty.");
            Directory.CreateDirectory(Path.GetDirectoryName(opt.ExcelOutputPath) ?? ".");
        }

        s.ListingWindow.WriteLine("Start...");

        int processed = 0, skipped = 0, errors = 0;
        int count = 0;
        string firstProcessedPath = null;

        foreach (var prt in Directory.EnumerateFiles(opt.RootFolder, "*.prt", SearchOption.AllDirectories))
        {
            // Фильтр по папкам
            if (!PathRules.PassesFolderFilter(prt, opt.IncludeFolderNames))
                continue;

            count++;

            var fi = new FileInfo(prt);
            string partNoFile = Path.GetFileNameWithoutExtension(prt);

            if (!repo.NeedsExtraction(prt, fi.Length, fi.LastWriteTimeUtc))
            {
                skipped++;
                continue;
            }

            try
            {
                var attrs = NxPartReader.ReadUserAttributesFromFile(s, prt);

                string designation = "";
                string v;
                if (attrs.TryGetValue("Обозначение", out v))
                    designation = v;

                string match = string.Equals(partNoFile, designation, StringComparison.OrdinalIgnoreCase)
                    ? "OK"
                    : "MISMATCH";

                if (firstProcessedPath == null)
                    firstProcessedPath = prt;

                repo.UpsertOk(prt, fi.Length, fi.LastWriteTimeUtc, partNoFile, designation, match, attrs);

                processed++;

                if (processed % 50 == 0)
                    s.ListingWindow.WriteLine($"Progress: processed={processed}, skipped={skipped}, errors={errors}, total_seen={count}");
            }
            catch (Exception exOne)
            {
                repo.UpsertError(prt, fi.Length, fi.LastWriteTimeUtc, partNoFile, exOne.Message);
                errors++;
            }
        }

        s.ListingWindow.WriteLine($"Scan done. processed={processed}, skipped={skipped}, errors={errors}, total_seen={count}");

        //// Диагностика (опционально)
        //if (firstProcessedPath != null)
        //{
        //    s.ListingWindow.WriteLine("DIAG Compare NX vs DB for: " + firstProcessedPath);

        //    //var nxAttrs = NxPartReader.ReadUserAttributesFromFile(s, firstProcessedPath);
        //    var dbAttrs = repo.GetAttributesByPath(firstProcessedPath);

        //    //s.ListingWindow.WriteLine($"NX attrs: {nxAttrs.Count}, DB attrs: {dbAttrs.Count}");
        //}

        // ===== Export (optional) =====
        if (opt.Mode == RunMode.ScanOnly)
        {
            s.ListingWindow.WriteLine("ScanOnly mode: export skipped.");
            return;
        }

        s.ListingWindow.WriteLine("Loading data from DB...");

        var attrNames = repo.GetAllAttributeNames();
        var parts = repo.GetAllParts(opt.RootFolder, opt.GroupSheets);

        // Фильтруем всегда, чтобы старые записи вне фильтра не попадали в отчёт
        parts = parts
            .FindAll(p => PathRules.PassesFolderFilter(p.FullPath, opt.IncludeFolderNames));

        for (int i = 0; i < parts.Count; i++)
        {
            parts[i].FolderSheetKey =
                PathRules.GetSheetKey(opt.RootFolder, parts[i].FullPath, opt.GroupSheets, opt.GroupMode);
        }

        s.ListingWindow.WriteLine("Exporting Excel...");
        ExcelExporter.Export(opt.ExcelOutputPath, parts, attrNames, opt.GroupSheets);
        s.ListingWindow.WriteLine("Excel saved: " + opt.ExcelOutputPath);
    }
}