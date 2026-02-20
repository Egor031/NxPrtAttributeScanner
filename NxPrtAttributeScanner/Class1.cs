using NXOpen;
using System;
using static NXOpen.NXObject;

public class NXEntryPoint
{
    public static void Main(string[] args)
    {
        Session s = Session.GetSession();
        s.ListingWindow.Open();

        Part workPart = s.Parts.Work;
        if (workPart == null)
        {
            s.ListingWindow.WriteLine("Нет открытой детали (Work Part == null). Открой PRT и запусти снова.");
            return;
        }

        s.ListingWindow.WriteLine($"WorkPart: {workPart.FullPath}");
        s.ListingWindow.WriteLine("User Attributes:");

        try
        {
            // В NX 1899 обычно это работает
            var attrs = workPart.GetUserAttributes();

            if (attrs == null || attrs.Length == 0)
            {
                s.ListingWindow.WriteLine("  (атрибутов нет)");
                return;
            }

            foreach (var a in attrs)
            {
                string title = (a.Title ?? "").Trim();
                if (title.Length == 0) continue;

                string value = AttributeToString(a);
                s.ListingWindow.WriteLine($"  {title} = {value}");
            }
        }
        catch (Exception ex)
        {
            s.ListingWindow.WriteLine("Ошибка чтения атрибутов: " + ex);
        }
    }

    private static string AttributeToString(AttributeInformation a)
    {
        switch (a.Type)
        {
            case NXObject.AttributeType.String:
                return a.StringValue ?? "";

            case NXObject.AttributeType.Integer:
                return a.IntegerValue.ToString();

            case NXObject.AttributeType.Real:
                return a.RealValue.ToString(System.Globalization.CultureInfo.InvariantCulture);

            case NXObject.AttributeType.Boolean:
                return a.BooleanValue ? "true" : "false";

            case NXObject.AttributeType.Time:
                return a.TimeValue ?? "";

            default:
                return "";
        }
    }

    public static int GetUnloadOption(string dummy)
    {
        return (int)Session.LibraryUnloadOption.Immediately;
    }
}