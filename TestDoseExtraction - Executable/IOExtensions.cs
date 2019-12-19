using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace System.IO
{
    static class IOExtensions
    {
        public static bool CheckForInvalidPathChars(string path)
        {
            if(path.Contains('*') ||
               path.Contains('?') ||
               path.Contains('"') ||
               path.Contains('<') ||
               path.Contains('>') ||
               path.Contains('|'))
            {
                return true;
            }
            return false;
        }

        // Reads in lines while ignoring blank/empty lines.
        public static string[] ReadNonBlankLines(string path)
        {
            string line;
            List<string> lines = new List<string>();

            using (StreamReader sr = new StreamReader(path))
                while (true)
                {
                    line = sr.ReadLine();

                    if (line == null)
                    {
                        break;
                    }
                    else if (line == Environment.NewLine || string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    lines.Add(line.Trim());
                }

            return lines.ToArray();
        }

        // Replaces invalid characters in a filename with underscores ('_').
        // Call this function on the file or directory name, NOT on the path.
        public static string MakeSafeFilename(string filename)
        {
            //char[] invalidChars = new char[] { '/', '?', '%', '*', ':', '|', '"', '<', '>', '.' };
            char[] invalidChars = new char[] { '/', '+', ':', ',', '.' };
            return string.Join("_", filename.Split(invalidChars));
        }
    }
}
