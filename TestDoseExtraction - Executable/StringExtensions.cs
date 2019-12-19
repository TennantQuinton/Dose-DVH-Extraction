using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StringExtensions
{
    // Extension methods must be defined in a static class.
    static class StringExtensions
    {
        // Removes characters from a string that are not letters.
        public static string RemoveNonLetters(this string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return str;
            }

            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        // Returns a substring from index startIndex of length maxLen or less, depending on
        // the number of characters in the substring.
        public static string GetSubstringByLength(this string str, int startIndex, int maxLen)
        {
            if (string.IsNullOrEmpty(str))
            {
                return str;
            }
            return str.Substring(startIndex, Math.Min(str.Length - startIndex, maxLen));
        }

        // Concatenates character c onto string str until the string length is equal to len.
        public static string CharFiller(this string str, char c, int len)
        {
            if (string.IsNullOrEmpty(str))
            {
                return str;
            }

            int strLen = str.Length;
            if (strLen >= len)
            {
                return str;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append(str);
            for (int i = strLen; i < len; i++)
            {
                sb.Append(c);
            }
            return sb.ToString();
        }
    }
}
