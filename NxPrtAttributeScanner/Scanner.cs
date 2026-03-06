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

        // ===============================
        // 1) Пытаемся построить карту: файл -> имя листа
        // ===============================
        var fileToSheet = BuildFileToSheetMap(s, opt, repo);

        bool useSheets = fileToSheet.Count > 0;
        if (useSheets)
            s.ListingWindow.WriteLine("Sheet mode: files will be taken only from folders assigned to Excel sheets.");
        else
            s.ListingWindow.WriteLine("Sheet mode: no configured sheet folders found, fallback to full Root scan.");

        int processed = 0, skipped = 0, errors = 0;
        int count = 0;
        string firstProcessedPath = null;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        IEnumerable<string> filesToScan;

        if (useSheets)
            filesToScan = fileToSheet.Keys;
        else
            filesToScan = Directory.EnumerateFiles(opt.RootFolder, "*.prt", SearchOption.AllDirectories);

        foreach (var prt in filesToScan)
        {
            seen.Add(prt);

            // В старом режиме оставляем старый фильтр
            if (!useSheets && !PathRules.PassesFolderFilter(prt, opt.IncludeFolderNames))
                continue;

            count++;
            var t0 = DateTime.Now;

            s.ListingWindow.WriteLine($"[{t0:HH:mm:ss}] Processing: {prt}");

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

        // Пока удаление оставляем только для старого режима.
        // Для sheet-mode нужен отдельный аккуратный механизм очистки по выбранным папкам.
        if (!useSheets)
        {
            if (seen.Count > 0)
            {
                repo.RemoveNotSeenUnderRoot(opt.RootFolder, seen);
            }
            else
            {
                s.ListingWindow.WriteLine("Внимание: не найдено ни одного файла по текущим фильтрам. База не очищалась.");
            }
        }
        else
        {
            s.ListingWindow.WriteLine("Sheet mode: cleanup of removed files is temporarily skipped in this version.");
        }

        s.ListingWindow.WriteLine($"Scan done. processed={processed}, skipped={skipped}, errors={errors}, total_seen={count}");

        // ===== Export (optional) =====
        if (opt.Mode == RunMode.ScanOnly)
        {
            s.ListingWindow.WriteLine("ScanOnly mode: export skipped.");
            return;
        }

        s.ListingWindow.WriteLine("Loading data from DB...");

        var attrNames = repo.GetAllAttributeNames();
        var parts = repo.GetAllParts(opt.RootFolder, true);

        if (useSheets)
        {
            // Берём только то, что относится к листам
            parts = parts.FindAll(p => fileToSheet.ContainsKey(p.FullPath));

            // Имя Excel-листа = имя листа из базы
            for (int i = 0; i < parts.Count; i++)
            {
                string sheetName;
                if (fileToSheet.TryGetValue(parts[i].FullPath, out sheetName))
                    parts[i].FolderSheetKey = sheetName;
                else
                    parts[i].FolderSheetKey = "UNASSIGNED";
            }
        }
        else
        {
            // Старое поведение
            parts = parts.FindAll(p => PathRules.PassesFolderFilter(p.FullPath, opt.IncludeFolderNames));

            for (int i = 0; i < parts.Count; i++)
            {
                parts[i].FolderSheetKey =
                    PathRules.GetSheetKey(opt.RootFolder, parts[i].FullPath, opt.GroupSheets, opt.GroupMode);
            }
        }

        s.ListingWindow.WriteLine("Exporting Excel...");
        if (parts.Count == 0)
        {
            s.ListingWindow.WriteLine("Нет данных для выгрузки в Excel (после фильтров). Excel не создан.");
            return;
        }

        // В sheet-mode всегда экспортируем по FolderSheetKey
        ExcelExporter.Export(opt.ExcelOutputPath, parts, attrNames, true);
        s.ListingWindow.WriteLine("Excel saved: " + opt.ExcelOutputPath);
    }

    private static Dictionary<string, string> BuildFileToSheetMap(Session s, ScanOptions opt, CacheRepository repo)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var sheets = repo.GetSheetsWithFolders();
        if (sheets == null || sheets.Count == 0)
            return map;

        foreach (var sheet in sheets)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(sheet.Name))
                continue;

            if (sheet.Folders == null || sheet.Folders.Count == 0)
                continue;

            foreach (var folder in sheet.Folders)
            {
                if (folder == null || string.IsNullOrWhiteSpace(folder.RelPath))
                    continue;

                string absFolder = CacheRepository.GetAbsolutePathFromRel(opt.RootFolder, folder.RelPath);

                if (string.IsNullOrWhiteSpace(absFolder) || !Directory.Exists(absFolder))
                {
                    s.ListingWindow.WriteLine($"[WARN] Sheet '{sheet.Name}': folder not found: {absFolder}");
                    continue;
                }

                var searchOption = folder.IncludeSubfolders
                    ? SearchOption.AllDirectories
                    : SearchOption.TopDirectoryOnly;

                IEnumerable<string> files;
                try
                {
                    files = Directory.EnumerateFiles(absFolder, "*.prt", searchOption);
                }
                catch (Exception ex)
                {
                    s.ListingWindow.WriteLine($"[WARN] Sheet '{sheet.Name}': cannot enumerate folder '{absFolder}': {ex.Message}");
                    continue;
                }

                foreach (var prt in files)
                {
                    // Старый фильтр includeFolderNames всё ещё применяем глобально
                    if (!PathRules.PassesFolderFilter(prt, opt.IncludeFolderNames))
                        continue;

                    string existingSheet;
                    if (map.TryGetValue(prt, out existingSheet))
                    {
                        // Если один и тот же файл попал в несколько листов — берём первый
                        if (!string.Equals(existingSheet, sheet.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            s.ListingWindow.WriteLine(
                                $"[WARN] File already assigned to sheet '{existingSheet}', skipped duplicate assignment to '{sheet.Name}': {prt}");
                        }
                        continue;
                    }

                    map[prt] = sheet.Name;
                }
            }
        }

        return map;
    }
}