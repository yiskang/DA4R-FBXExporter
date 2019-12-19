using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Autodesk.ADN.FbxExporter
{
    static class StringExtension
    {
        public static string ReplaceInvalidFileNameChars(this string s, string replacement = "")
        {
            return Regex.Replace(s,
              "[" + Regex.Escape(new String(System.IO.Path.GetInvalidPathChars())) + "]",
              replacement, //can even use a replacement string of any length
              RegexOptions.IgnoreCase);
            //not using System.IO.Path.InvalidPathChars (deprecated insecure API)
        }
    }
}
