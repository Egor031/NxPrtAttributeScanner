// RunScanner.cs - loader journal for run_journal.exe
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
            // 1) Определяем baseDir:
            //    - если передан config=... -> папка конфига
            //    - иначе journalPath=... -> папка journal
            //    - иначе текущая директория
            string baseDir = ResolveBaseDir(args);

            // 2) Где лежит DLL сканера?
            string dllPath = GetArgValue(args, "scannerDll");
            if (string.IsNullOrWhiteSpace(dllPath))
            {
                dllPath = Path.Combine(baseDir, "NxPrtAttributeScanner.dll");
            }
            else
            {
                dllPath = Environment.ExpandEnvironmentVariables(dllPath);
                if (!Path.IsPathRooted(dllPath))
                    dllPath = Path.GetFullPath(Path.Combine(baseDir, dllPath));
                else
                    dllPath = Path.GetFullPath(dllPath);
            }

            if (!File.Exists(dllPath))
                throw new FileNotFoundException("Scanner DLL not found: " + dllPath);

            // 3) Подгружаем сборку и вызываем HeadlessEntryPoint.Main(args)
            var asm = Assembly.LoadFrom(dllPath);

            var t = asm.GetType("HeadlessEntryPoint", throwOnError: true);
            var m = t.GetMethod("Main", BindingFlags.Public | BindingFlags.Static);
            if (m == null)
                throw new MissingMethodException("HeadlessEntryPoint.Main not found in " + dllPath);

            m.Invoke(null, new object[] { args });
        }
        catch (TargetInvocationException tie)
        {
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

    private static string ResolveBaseDir(string[] args)
    {
        // config=... (самый надёжный якорь)
        string configPath = GetArgValue(args, "config");
        if (!string.IsNullOrWhiteSpace(configPath))
        {
            configPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(configPath));
            string dir = Path.GetDirectoryName(configPath);
            if (!string.IsNullOrWhiteSpace(dir))
                return dir;
        }

        // journalPath=... (опционально)
        string journalPath = GetArgValue(args, "journalPath");
        if (!string.IsNullOrWhiteSpace(journalPath))
        {
            journalPath = Path.GetFullPath(Environment.ExpandEnvironmentVariables(journalPath));
            string dir = Path.GetDirectoryName(journalPath);
            if (!string.IsNullOrWhiteSpace(dir))
                return dir;
        }

        // fallback
        return Directory.GetCurrentDirectory();
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