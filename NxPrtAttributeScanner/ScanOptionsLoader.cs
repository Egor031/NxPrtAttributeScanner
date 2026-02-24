using System;
using System.IO;

public static class ScanOptionsLoader
{
    public static ScanOptions LoadFromIni(string iniPath)
    {
        if (string.IsNullOrWhiteSpace(iniPath))
            throw new Exception("iniPath is empty.");

        if (!File.Exists(iniPath))
            throw new FileNotFoundException("Config ini not found: " + iniPath);

        var ini = IniFile.Load(iniPath);

        // База для portable: папка, где лежит scan.ini
        string baseDir = Path.GetDirectoryName(Path.GetFullPath(iniPath)) ?? Environment.CurrentDirectory;

        var opt = new ScanOptions();

        // Root (можно относительный)
        opt.RootFolder = ini.GetString("Scan", "Root", "").Trim();
        if (string.IsNullOrWhiteSpace(opt.RootFolder))
            throw new Exception("Scan.Root is empty in ini.");

        opt.RootFolder = MakeAbsolute(baseDir, opt.RootFolder);

        // Excel output (можно относительный)
        string outPath = ini.GetString("Excel", "Out", @".\reports\Parts.xlsx").Trim();
        opt.ExcelOutputPath = MakeAbsolute(baseDir, outPath);

        opt.GroupSheets = ini.GetBool01("Excel", "GroupSheets", true);

        string gm = (ini.GetString("Excel", "GroupMode", "FirstLevel") ?? "").Trim();
        opt.GroupMode = gm.Equals("AllInOne", StringComparison.OrdinalIgnoreCase)
            ? SheetGroupingMode.AllInOne
            : SheetGroupingMode.FirstLevel;

        string mode = (ini.GetString("Scan", "Mode", "ScanAndExport") ?? "").Trim();
        opt.Mode = mode.Equals("ScanOnly", StringComparison.OrdinalIgnoreCase)
            ? RunMode.ScanOnly
            : RunMode.ScanAndExport;

        opt.IncludeFolderNames = ini.GetStrings("Scan", "IncludeFolderName");
        // Если IniFile.GetStrings не триммит — можно дотриммить тут:
        for (int i = opt.IncludeFolderNames.Count - 1; i >= 0; i--)
        {
            var t = (opt.IncludeFolderNames[i] ?? "").Trim();
            if (t.Length == 0) opt.IncludeFolderNames.RemoveAt(i);
            else opt.IncludeFolderNames[i] = t;
        }

        // DbPath (можно относительный)
        string dbPath = ini.GetString("Scan", "DbPath", @".\cache\parts.db").Trim();
        opt.DbPath = MakeAbsolute(baseDir, dbPath);

        return opt;
    }

    private static string MakeAbsolute(string baseDir, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return path;

        // Разрешим %ENV% переменные (приятно для пользователей)
        path = Environment.ExpandEnvironmentVariables(path);

        if (Path.IsPathRooted(path))
            return Path.GetFullPath(path);

        return Path.GetFullPath(Path.Combine(baseDir, path));
    }
}