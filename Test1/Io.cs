using System;
using System.IO;

namespace Test1
{
    static class Io
    {
        internal static string Desktopize(params string[] paths) => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), Path.Combine(paths));
    }
}
