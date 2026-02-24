using NXOpen;
using System;
using System.IO;

public class RunScanner
{
    public static void Main(string[] args)
    {
        Session s = Session.GetSession();
        s.ListingWindow.Open();

        try
        {
            // args: config=path_to_ini
            string iniPath = GetArgValue(args, "config");
            if (string.IsNullOrWhiteSpace(iniPath))
                throw new Exception("Missing argument: config=<path_to_scan.ini>");

            var opt = ScanOptionsLoader.LoadFromIni(iniPath);
            var repo = new CacheRepository(opt.DbPath);

            Scanner.Run(s, opt, repo);
        }
        catch (Exception ex)
        {
            try
            {
                s.ListingWindow.WriteLine("FATAL ERROR:");
                s.ListingWindow.WriteLine(ex.ToString());
            }
            catch { }

            try
            {
                File.WriteAllText(Path.Combine(Environment.CurrentDirectory, "last_error.txt"), ex.ToString());
            }
            catch { }
        }
    }

    private static string GetArgValue(string[] args, string key)
    {
        // ожидаем формат: key=value
        if (args == null) return null;
        for (int i = 0; i < args.Length; i++)
        {
            var a = args[i];
            if (a == null) continue;
            int eq = a.IndexOf('=');
            if (eq <= 0) continue;

            string k = a.Substring(0, eq).Trim();
            if (!k.Equals(key, StringComparison.OrdinalIgnoreCase)) continue;

            return a.Substring(eq + 1).Trim().Trim('"');
        }
        return null;
    }

    public static int GetUnloadOption(string dummy)
    {
        return (int)Session.LibraryUnloadOption.Immediately;
    }
}