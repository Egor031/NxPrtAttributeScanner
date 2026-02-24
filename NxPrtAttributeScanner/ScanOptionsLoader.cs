using System;
using System.Collections.Generic;
using System.IO;

public static class ScanOptionsLoader
{
    public static ScanOptions LoadFromIni(string iniPath)
    {
        if (!File.Exists(iniPath))
            throw new FileNotFoundException("Config ini not found: " + iniPath);

        var ini = IniFile.Load(iniPath);

        var opt = new ScanOptions();

        opt.RootFolder = ini.GetString("Scan", "Root", "");
        if (string.IsNullOrWhiteSpace(opt.RootFolder))
            throw new Exception("Scan.Root is empty in ini.");

        opt.ExcelOutputPath = ini.GetString("Excel", "Out",
            Path.Combine(Environment.CurrentDirectory, "Parts.xlsx"));

        opt.GroupSheets = ini.GetBool01("Excel", "GroupSheets", true);

        string gm = ini.GetString("Excel", "GroupMode", "FirstLevel");
        opt.GroupMode = gm.Equals("AllInOne", StringComparison.OrdinalIgnoreCase)
            ? SheetGroupingMode.AllInOne
            : SheetGroupingMode.FirstLevel;

        string mode = ini.GetString("Scan", "Mode", "ScanAndExport");
        opt.Mode = mode.Equals("ScanOnly", StringComparison.OrdinalIgnoreCase)
            ? RunMode.ScanOnly
            : RunMode.ScanAndExport;

        opt.IncludeFolderNames = ini.GetStrings("Scan", "IncludeFolderName");

        // DbPath отдельным полем не нужен в ScanOptions — он нужен для CacheRepository,
        // но оставим чтение тут для удобства лаунчера:
        opt.DbPath = ini.GetString("Scan", "DbPath",
            Path.Combine(Environment.CurrentDirectory, "cache", "parts.db"));

        return opt;
    }
}