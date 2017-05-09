using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

public static class RegexUtils
{

    public static string Repeat(this string value, int count)
    {
        return new StringBuilder(value.Length * count).Insert(0, value, count).ToString();
    }

    public static bool IsEmpty(this string input)
    {
        return string.IsNullOrWhiteSpace(input);
    }

    public static bool matches(this string input, string pattern)
    {
        return Regex.IsMatch(input, pattern);
    }

    public static string ReplaceFirst(this string input, string pattern, string replacement)
    {
        return new Regex(pattern).Replace(input, replacement, 1);
    }
}

