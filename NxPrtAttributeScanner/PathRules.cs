using System;
using System.Collections.Generic;
using System.IO;

public static class PathRules
{
    public static bool PassesFolderFilter(string fullPath, List<string> includeFolderNames)
    {
        // Пусто => без фильтра
        if (includeFolderNames == null || includeFolderNames.Count == 0)
            return true;

        // Разбираем путь на имена папок и сравниваем по имени папки (case-insensitive)
        // fullPath может быть очень длинным — делаем без лишних аллокаций насколько возможно.
        string dir = Path.GetDirectoryName(fullPath) ?? "";
        if (dir.Length == 0) return false;

        // Пройдём по сегментам пути
        // Split простой и надёжный; если захочешь микрооптимизацию — сделаем позже.
        var parts = dir.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        for (int i = 0; i < parts.Length; i++)
        {
            string folder = parts[i];
            if (string.IsNullOrWhiteSpace(folder)) continue;

            for (int j = 0; j < includeFolderNames.Count; j++)
            {
                if (string.Equals(folder, includeFolderNames[j], StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }

        return false;
    }

    public static string GetSheetKey(string rootFolder, string fullPath, bool groupSheets, SheetGroupingMode mode)
    {
        if (!groupSheets || mode == SheetGroupingMode.AllInOne)
            return "ALL";

        // FirstLevel: первая подпапка относительно root
        string rel = GetRelativePathSafe(rootFolder, fullPath);
        if (string.IsNullOrEmpty(rel))
            return "ROOT";

        int idx = rel.IndexOf(Path.DirectorySeparatorChar);
        if (idx < 0)
            return "ROOT"; // файл прямо в root

        string first = rel.Substring(0, idx);
        return string.IsNullOrWhiteSpace(first) ? "ROOT" : first;
    }

    // .NET Framework: нет Path.GetRelativePath, делаем через Uri
    public static string GetRelativePathSafe(string rootFolder, string fullPath)
    {
        if (string.IsNullOrEmpty(rootFolder) || string.IsNullOrEmpty(fullPath))
            return "";

        string root = rootFolder;
        if (!root.EndsWith(Path.DirectorySeparatorChar.ToString()))
            root += Path.DirectorySeparatorChar;

        try
        {
            var rootUri = new Uri(root);
            var fileUri = new Uri(fullPath);

            if (rootUri.Scheme != fileUri.Scheme)
                return "";

            var relUri = rootUri.MakeRelativeUri(fileUri);
            string rel = Uri.UnescapeDataString(relUri.ToString());
            rel = rel.Replace('/', Path.DirectorySeparatorChar);
            return rel;
        }
        catch
        {
            return "";
        }
    }
}