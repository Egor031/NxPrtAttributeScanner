using NXOpen;
using System;
using System.Collections.Generic;
using System.IO;
using static NXOpen.NXObject;

public class NXEntryPoint
{
    public static void Main(string[] args)
    {
        var s = Session.GetSession();
        s.ListingWindow.Open();

        string root = @"D:\ZherlitsynEE\SaveFormatTest\Test\PRT"; // <-- папка с .prt

        int count = 0;
        foreach (var prt in Directory.EnumerateFiles(root, "*.prt", SearchOption.AllDirectories))
        {
            count++;

            string partNoFile = Path.GetFileNameWithoutExtension(prt);

            try
            {
                var attrs = NxPartReader.ReadUserAttributesFromFile(s, prt);

                string designation = attrs.TryGetValue("Обозначение", out var v) ? v : "";
                string match = string.Equals(partNoFile, designation, StringComparison.OrdinalIgnoreCase) ? "OK" : "MISMATCH";

                s.ListingWindow.WriteLine($"{count}. {partNoFile} | Обозначение={designation} | {match}");
            }
            catch (Exception ex)
            {
                s.ListingWindow.WriteLine($"{count}. {prt} | ERROR: {ex.Message}");
            }

            // Чтобы не улететь в бесконечность на тесте:
            if (count >= 50) break;
        }

        s.ListingWindow.WriteLine("Done.");
    }

    public static int GetUnloadOption(string dummy)
        => (int)Session.LibraryUnloadOption.Immediately;
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

    private static string AttributeToString(AttributeInformation a)
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