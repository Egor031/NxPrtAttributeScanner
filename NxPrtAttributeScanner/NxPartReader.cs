using NXOpen;
using System;
using System.Collections.Generic;

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

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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
            try { if (pls != null) pls.Dispose(); } catch { }

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
            case NXOpen.NXObject.AttributeType.String: return a.StringValue ?? "";
            case NXOpen.NXObject.AttributeType.Integer: return a.IntegerValue.ToString();
            case NXOpen.NXObject.AttributeType.Real: return a.RealValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
            case NXOpen.NXObject.AttributeType.Boolean: return a.BooleanValue ? "true" : "false";
            case NXOpen.NXObject.AttributeType.Time: return a.TimeValue ?? "";
            default: return "";
        }
    }
}