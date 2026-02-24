using NXOpen;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using static NXOpen.CAE.Post;


public class NXEntryPoint
{
    public static void Main(string[] args)
    {
        Session s = Session.GetSession();
        s.ListingWindow.Open();

        try
        {
            string root = @"D:\ZherlitsynEE\SaveFormatTest\Test\PRT";

            var opt = new ScanOptions
            {
                RootFolder = root,
                GroupSheets = true,
                GroupMode = SheetGroupingMode.FirstLevel,
                ExcelOutputPath = @"D:\ZherlitsynEE\NxPrtAttributeScanner\reports\Parts.xlsx"
            };

            // Пример фильтра
            opt.IncludeFolderNames.Add("Кронштейны");
            opt.Mode = RunMode.ScanAndExport;
            //opt.Mode = RunMode.ScanOnly;

            var repo = new CacheRepository(
                @"D:\ZherlitsynEE\NxPrtAttributeScanner\cache\parts.db");

            Scanner.Run(s, opt, repo);
        }
        catch (Exception ex)
        {
            s.ListingWindow.WriteLine("FATAL ERROR:");
            s.ListingWindow.WriteLine(ex.ToString());
        }
    }

    public static int GetUnloadOption(string dummy)
    {
        return (int)Session.LibraryUnloadOption.Immediately;
    }
}

public static class NxPartReader
{
    public static Dictionary<string, string> ReadUserAttributesFromFile(Session s, string prtPath)
    {
        BasePart basePart = null;
        PartLoadStatus pls = null;

        try
        {
            basePart = s.Parts.OpenBaseDisplay(prtPath, out pls);

            var part = basePart as Part;
            if (part == null)
                throw new Exception("Открытый файл не является Part.");

            var dict = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);

            var attrs = part.GetUserAttributes();
            foreach (var a in attrs)
            {
                var title = (a.Title ?? "").Trim();
                if (title.Length == 0) continue;

                dict[title] = AttributeToString(a);
            }

            return dict;
        }
        finally
        {
            try { pls?.Dispose(); } catch { }

            try
            {
                if (basePart != null)
                    basePart.Close(BasePart.CloseWholeTree.False, BasePart.CloseModified.CloseModified, null);
            }
            catch { }
        }
    }

    private static string AttributeToString(NXOpen.NXObject.AttributeInformation a)
    {
        switch (a.Type)
        {
            case NXObject.AttributeType.String: return a.StringValue ?? "";
            case NXObject.AttributeType.Integer: return a.IntegerValue.ToString();
            case NXObject.AttributeType.Real: return a.RealValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
            case NXObject.AttributeType.Boolean: return a.BooleanValue ? "true" : "false";
            case NXObject.AttributeType.Time: return a.TimeValue ?? "";
            default: return "";
        }
    }
}