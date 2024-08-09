using System.Collections.Generic;
using System.Linq;

public static class StringExtensions
{
    public static string RemoveToneMarks(this string pinyin)
    {
        if (string.IsNullOrWhiteSpace(pinyin))
        {
            return pinyin;
        }

        var toneMap = new Dictionary<char, char>
        {
            { 'ā', 'a' }, { 'á', 'a' }, { 'ǎ', 'a' }, { 'à', 'a' },
            { 'ē', 'e' }, { 'é', 'e' }, { 'ě', 'e' }, { 'è', 'e' },
            { 'ī', 'i' }, { 'í', 'i' }, { 'ǐ', 'i' }, { 'ì', 'i' },
            { 'ō', 'o' }, { 'ó', 'o' }, { 'ǒ', 'o' }, { 'ò', 'o' },
            { 'ū', 'u' }, { 'ú', 'u' }, { 'ǔ', 'u' }, { 'ù', 'u' },
            { 'ǖ', 'ü' }, { 'ǘ', 'ü' }, { 'ǚ', 'ü' }, { 'ǜ', 'ü' }
        };

        return new string(pinyin.Select(c => toneMap.ContainsKey(c) ? toneMap[c] : c).ToArray());
    }
}
