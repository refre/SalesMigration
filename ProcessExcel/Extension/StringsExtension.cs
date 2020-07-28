using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ista.Migration.Excel
{
    public static class StringExtensions
    {
        public static bool StartsWithAny(this string str,params string[] values)
        {
            if (!string.IsNullOrEmpty(str) || values.Length > 0)
            {
                foreach (string value in values)
                {
                    if (str.StartsWith(value))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
