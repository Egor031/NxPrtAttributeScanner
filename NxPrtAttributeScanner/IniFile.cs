using System;
using System.Collections.Generic;
using System.IO;

public sealed class IniFile
{
    private readonly Dictionary<string, Dictionary<string, List<string>>> _data;

    public IniFile()
    {
        _data = new Dictionary<string, Dictionary<string, List<string>>>(StringComparer.OrdinalIgnoreCase);
    }

    public static IniFile Load(string path)
    {
        var ini = new IniFile();
        ini.Read(path);
        return ini;
    }

    private void Read(string path)
    {
        string section = "";

        foreach (var rawLine in File.ReadAllLines(path))
        {
            var line = rawLine.Trim();
            if (line.Length == 0) continue;
            if (line.StartsWith(";") || line.StartsWith("#")) continue;

            if (line.StartsWith("[") && line.EndsWith("]"))
            {
                section = line.Substring(1, line.Length - 2).Trim();
                if (!_data.ContainsKey(section))
                    _data[section] = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                continue;
            }

            int eq = line.IndexOf('=');
            if (eq <= 0) continue;

            string key = line.Substring(0, eq).Trim();
            string val = line.Substring(eq + 1).Trim();

            Dictionary<string, List<string>> sec;
            if (!_data.TryGetValue(section, out sec))
            {
                sec = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                _data[section] = sec;
            }

            List<string> list;
            if (!sec.TryGetValue(key, out list))
            {
                list = new List<string>();
                sec[key] = list;
            }

            list.Add(val);
        }
    }

    public string GetString(string section, string key, string defaultValue)
    {
        List<string> vals;
        if (_data.ContainsKey(section) && _data[section].TryGetValue(key, out vals) && vals.Count > 0)
            return vals[0];
        return defaultValue;
    }

    public List<string> GetStrings(string section, string key)
    {
        List<string> vals;
        if (_data.ContainsKey(section) && _data[section].TryGetValue(key, out vals))
            return new List<string>(vals);
        return new List<string>();
    }

    public bool GetBool01(string section, string key, bool defaultValue)
    {
        string s = GetString(section, key, defaultValue ? "1" : "0");
        return s == "1" || s.Equals("true", StringComparison.OrdinalIgnoreCase) || s.Equals("yes", StringComparison.OrdinalIgnoreCase);
    }
}