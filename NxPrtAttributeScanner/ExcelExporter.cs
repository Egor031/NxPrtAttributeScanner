using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Table;

public static class ExcelExporter
{
    public static void Export(string outputXlsxPath, List<PartRow> parts, List<string> attrNames, bool groupByFolderSheets)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputXlsxPath) ?? ".");

        // EPPlus 4.x: лицензии/LicenseContext нет
        using (var pkg = new ExcelPackage())
        {
            var groups = groupByFolderSheets
                ? parts.GroupBy(p => p.FolderSheetKey, StringComparer.OrdinalIgnoreCase)
                : new[] { parts.GroupBy(p => "ALL").First() };

            foreach (var g in groups)
            {
                string sheetName = MakeSafeSheetName(g.Key);
                sheetName = EnsureUniqueSheetName(pkg, sheetName);

                var ws = pkg.Workbook.Worksheets.Add(sheetName);

                // Заголовки
                var headers = new List<string>
                {
                    "PartNo_File",
                    "Designation_Attr",
                    "Match",
                    "FullPath",
                    "LastWriteTimeUtc",
                    "ExtractedAtUtc"
                };
                headers.AddRange(attrNames);

                for (int c = 0; c < headers.Count; c++)
                    ws.Cells[1, c + 1].Value = headers[c];

                // Данные
                int row = 2;
                foreach (var p in g)
                {
                    int col = 1;
                    ws.Cells[row, col++].Value = p.PartNoFile;
                    ws.Cells[row, col++].Value = p.Designation;
                    ws.Cells[row, col++].Value = p.Match;
                    ws.Cells[row, col++].Value = p.FullPath;
                    ws.Cells[row, col++].Value = p.LastWriteUtc.ToString("O");
                    ws.Cells[row, col++].Value = p.ExtractedUtc.ToString("O");

                    for (int i = 0; i < attrNames.Count; i++)
                    {
                        string v;
                        p.Attrs.TryGetValue(attrNames[i], out v);
                        ws.Cells[row, col++].Value = v ?? "";
                    }

                    row++;
                }

                int lastRow = Math.Max(1, row - 1);
                int lastCol = headers.Count;

                // Таблица + фильтры
                var range = ws.Cells[1, 1, lastRow, lastCol];
                var table = ws.Tables.Add(range, "T_" + SanitizeTableName(sheetName));
                table.TableStyle = TableStyles.Medium2;
                table.ShowFilter = true;

                // Автоширина
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
            }

            // Сохранение
            var fi = new FileInfo(outputXlsxPath);
            if (fi.Exists) fi.Delete();
            pkg.SaveAs(fi);
        }
    }

    private static string MakeSafeSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            name = "Sheet";

        // Excel запрещает: : \ / ? * [ ]
        char[] bad = { ':', '\\', '/', '?', '*', '[', ']' };
        foreach (var c in bad)
            name = name.Replace(c, '_');

        name = name.Trim().Trim('\'');

        if (name.Length > 31)
            name = name.Substring(0, 31);

        if (string.IsNullOrWhiteSpace(name))
            name = "Sheet";

        return name;
    }

    private static string EnsureUniqueSheetName(ExcelPackage pkg, string baseName)
    {
        string name = baseName;
        int i = 2;

        while (pkg.Workbook.Worksheets.Any(w => string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            string suffix = "_" + i.ToString();
            int maxBase = Math.Max(1, 31 - suffix.Length);
            string cut = baseName.Length > maxBase ? baseName.Substring(0, maxBase) : baseName;
            name = cut + suffix;
            i++;
        }

        return name;
    }

    private static string SanitizeTableName(string sheetName)
    {
        // EPPlus имя таблицы: буквы/цифры/подчёркивание, не начинать с цифры
        var s = new string(sheetName.Select(ch =>
            char.IsLetterOrDigit(ch) ? ch : '_'
        ).ToArray());

        if (string.IsNullOrWhiteSpace(s)) s = "Table";
        if (char.IsDigit(s[0])) s = "T" + s;
        return s.Length > 25 ? s.Substring(0, 25) : s;
    }
}