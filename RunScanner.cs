// RunScanner.cs - tiny loader journal for run_journal.exe
using System;
using System.IO;
using System.Reflection;
using NXOpen;

public class RunScanner
{
    public static void Main(string[] args)
    {
        Session s = Session.GetSession();
        s.ListingWindow.Open();

        try
        {
            // 1) Где лежит DLL сканера?
            // По умолчанию: рядом с этим .cs
            string baseDir = Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
            // baseDir тут будет папка NX, поэтому лучше определять по самому journal-файлу:
            string journalPath = GetArgValue(args, "journalPath"); // опционально
            string runDir = !string.IsNullOrWhiteSpace(journalPath)
                ? Path.GetDirectoryName(journalPath)
                : Directory.GetCurrentDirectory();

            // Можно передать scannerDll=... явно, если хочешь
            string dllPath = GetArgValue(args, "scannerDll");
            if (string.IsNullOrWhiteSpace(dllPath))
                dllPath = Path.Combine(runDir, "NxPrtAttributeScanner.dll");

            if (!File.Exists(dllPath))
                throw new FileNotFoundException("Scanner DLL not found: " + dllPath);

            // 2) Подгружаем сборку и вызываем HeadlessEntryPoint.Main(config=...)
            var asm = Assembly.LoadFrom(dllPath);

            var t = asm.GetType("HeadlessEntryPoint", throwOnError: true);
            var m = t.GetMethod("Main", BindingFlags.Public | BindingFlags.Static);
            if (m == null)
                throw new MissingMethodException("HeadlessEntryPoint.Main not found in " + dllPath);

            // Передаём args как есть (там будет config=... и т.п.)
            m.Invoke(null, new object[] { args });
        }
        catch (TargetInvocationException tie)
        {
            // если внутри твоего кода упало исключение — покажем настоящее
            var ex = tie.InnerException ?? tie;
            s.ListingWindow.WriteLine("FATAL ERROR (inner):");
            s.ListingWindow.WriteLine(ex.ToString());
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
        for (int i = 0; i < args.Length; i++)
        {
            var a = args[i];
            if (string.IsNullOrWhiteSpace(a)) continue;
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
        return (int)NXOpen.Session.LibraryUnloadOption.Immediately;
    }
}