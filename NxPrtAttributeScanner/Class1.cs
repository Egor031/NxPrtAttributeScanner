using NXOpen;
using System;
using System.Collections.Generic;
using System.IO;


public class NXEntryPoint
{
    public static void Main(string[] args)
    {
        Session s = Session.GetSession();
        s.ListingWindow.Open();

        try
        {
            s.ListingWindow.WriteLine("Start...");

            string root = @"D:\ZherlitsynEE\SaveFormatTest\Test\PRT";
            var repo = new CacheRepository(@"D:\ZherlitsynEE\NxPrtAttributeScanner\cache\parts.db");

            int processed = 0, skipped = 0, errors = 0;
            int count = 0;

            foreach (var prt in Directory.EnumerateFiles(root, "*.prt", SearchOption.AllDirectories))
            {
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
                    if (attrs.TryGetValue("Обозначение", out v)) designation = v;

                    string match = string.Equals(partNoFile, designation, StringComparison.OrdinalIgnoreCase) ? "OK" : "MISMATCH";

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

            // ===== Excel export =====
            bool groupByFolderSheets = true;

            s.ListingWindow.WriteLine("Loading data from DB...");
            var attrNames = repo.GetAllAttributeNames();
            var parts = repo.GetAllParts(root, groupByFolderSheets);

            string outXlsx = @"D:\ZherlitsynEE\NxPrtAttributeScanner\reports\Parts.xlsx";
            s.ListingWindow.WriteLine("Exporting Excel to: " + outXlsx);

            ExcelExporter.Export(outXlsx, parts, attrNames, groupByFolderSheets);

            s.ListingWindow.WriteLine("Excel saved: " + outXlsx);
        }
        catch (Exception ex)
        {
            // Главное: увидеть реальную причину
            s.ListingWindow.WriteLine("FATAL ERROR:");
            s.ListingWindow.WriteLine(ex.ToString());

            // И на всякий случай в файл
            try
            {
                File.WriteAllText(@"D:\ZherlitsynEE\NxPrtAttributeScanner\reports\last_error.txt", ex.ToString());
                s.ListingWindow.WriteLine("Error log saved to last_error.txt");
            }
            catch { }
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