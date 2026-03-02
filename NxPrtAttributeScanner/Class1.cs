// Demo entrypoint для запуска прямо из NX (не для production)
// Настройки берутся из аргументов или из scan.ini рядом с dll.

using NXOpen;
using System;
using System.Collections.Generic;
using System.IO;

public class NXEntryPoint
{
    public static void Main(string[] args)
    {
        var s = Session.GetSession();
        s.ListingWindow.Open();

        try
        {
            // 1) Если передан config=... — используем ini
            string ini = GetArgValue(args, "config");
            ScanOptions opt;

            if (!string.IsNullOrWhiteSpace(ini))
            {
                ini = Path.GetFullPath(Environment.ExpandEnvironmentVariables(ini));
                opt = ScanOptionsLoader.LoadFromIni(ini);
            }
            else
            {
                // 2) Иначе — минимальные дефолты относительно папки программы
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                opt = new ScanOptions
                {
                    RootFolder = GetArgValue(args, "root") ?? "",
                    ExcelOutputPath = Path.Combine(baseDir, "reports", "Parts.xlsx"),
                    Mode = RunMode.ScanAndExport,
                    GroupMode = SheetGroupingMode.FirstLevel,
                    GroupSheets = true,
                    DbPath = Path.Combine(baseDir, "cache", "parts.db"),
                };

                if (string.IsNullOrWhiteSpace(opt.RootFolder))
                    throw new Exception("Не задан root. Передай аргумент root=<папка> или config=<scan.ini>");
            }

            var repo = new CacheRepository(opt.DbPath);
            Scanner.Run(s, opt, repo);
        }
        catch (Exception ex)
        {
            s.ListingWindow.WriteLine("FATAL ERROR:");
            s.ListingWindow.WriteLine(ex.ToString());
        }
    }

    private static string GetArgValue(string[] args, string key)
    {
        if (args == null) return null;
        foreach (var a in args)
        {
            if (string.IsNullOrWhiteSpace(a)) continue;
            int eq = a.IndexOf('=');
            if (eq <= 0) continue;
            var k = a.Substring(0, eq).Trim();
            if (!k.Equals(key, StringComparison.OrdinalIgnoreCase)) continue;
            return a.Substring(eq + 1).Trim().Trim('"');
        }
        return null;
    }

    public static int GetUnloadOption(string dummy)
        => (int)Session.LibraryUnloadOption.Immediately;
}