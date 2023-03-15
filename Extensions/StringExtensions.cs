using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace FT_ADDON
{
    static class StringExtensions
    {
        const string space = " ";

        // Add space between uppercase and lowercase
        public static string NaturalSpacing(this string str)
        {
            StringBuilder sb = new StringBuilder();
            return str.Select(c => sb.Append(char.IsLower(c) || sb.Length == 0 ? c.ToString() : space + c.ToString())).Last().ToString();
        }

        public static string FirstCharToUpper(this string input)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(input.ToLower());
        }
    }
}
