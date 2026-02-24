using System;
using System.Collections.Generic;

public enum SheetGroupingMode
{
    AllInOne,
    FirstLevel
}
public enum RunMode
{
    ScanAndExport,
    ScanOnly
}

public sealed class ScanOptions
{
    public string RootFolder;
    public List<string> IncludeFolderNames = new List<string>(); // "Кронштейны", "Токарная", ...
    public bool GroupSheets;
    public SheetGroupingMode GroupMode = SheetGroupingMode.FirstLevel;
    public string ExcelOutputPath;
    public RunMode Mode = RunMode.ScanAndExport;
    public string DbPath;
}
