using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PanelAddinWizard
{
    static class Translit
    {
        static Dictionary<Char, string> _translitMap = new Dictionary<char, string>();

        public static string MakeIdentifierFromString(string name)
        {
            var result = new StringBuilder();

            if (Char.IsDigit(name[0]))
                result.Append("_");

            for (int i = 0; i < name.Length; ++i)
                result.Append(TranslitChar(name[i]));

            return result.ToString();
        }

        const string _validChars =
            "._1234567890qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM";

        const string _stringSrc =
            "а|б|в|г|д|е|ё|ж|з|и|й|к|л|м|н|о|п|р|с|т|у|ф|х|ц|ч|ш|щ|ъ|ы|ь|э|ю|я|А|Б|В|Г|Д|Е|Ё|Ж|З|И|Й|К|Л|М|Н|О|П|Р|С|Т|У|Ф|Х|Ц|Ч|Ш|Щ|Ъ|Ы|Ь|Э|Ю|Я|" +
            "á|ä|č|ď|é|ě|í|ľ|ĺ|ň|ó|ô|ř|ŕ|š|ť|ú|ů|ý|ž|Á|Ä|Č|Ď|É|Ě|Í|Ľ|Ĺ|Ň|Ó|Ô|Ř|Ŕ|Š|Ť|Ú|Ů|Ý|Ž|ß|" +
            "À|Á|Â|Ã|Ä|Å|" +
            "È|É|Ê|Ë|" +
            "Ì|Í|Î|Ï|" +
            "Ò|Ó|Ô|Õ|Ö|" +
            "Ù|Ú|Û|Ü|" +
            "Ý|Ç|Ć|Ĉ|" +
            "à|á|â|ã|ä|å|" +
            "è|é|ê|ë|" +
            "ì|í|î|ï|" +
            "ò|ó|ô|õ|ö|" +
            "ù|ú|û|ü|" +
            "ý|ç|ć|ĉ";

        const string _stringDst =
            "a|b|v|g|d|e|yo|zh|z|i|j|k|l|m|n|o|p|r|s|t|u|f|h|c|ch|sh|sh||y||e|yu|ya|A|B|V|G|D|E|Yo|Zh|Z|I|J|K|L|M|N|O|P|R|S|T|U|F|H|C|Ch|Sh|Sh||Y||E|Yu|Ya|" +
            "a|a|c|d|e|e|i|l|l|n|o|o|r|r|s|t|u|u|y|z|A|A|C|D|E|E|I|L|L|N|O|O|R|R|S|T|U|U|Y|Z|ss|" +
            "A|A|A|A|A|A|" +
            "E|E|E|E|" +
            "I|I|I|I|" +
            "O|O|O|O|O|" +
            "U|U|U|U|" +
            "Y|C|C|C|" +
            "a|a|a|a|a|a|" +
            "e|e|e|e|" +
            "i|i|i|i|" +
            "o|o|o|o|o|" +
            "u|u|u|u|" +
            "y|c|c|c";

        static string TranslitChar(Char ch)
        {
            if (_validChars.IndexOf(ch) >= 0)
                return ch.ToString();

            string result;
            if (_translitMap.TryGetValue(ch, out result))
                return result;

			return "_";
        }

        static Translit()
        {
            var src = _stringSrc.Split('|');
            var dst = _stringDst.Split('|');

            for (int i = 0; i < src.Length; ++i)
                _translitMap[src[i][0]] = dst[i];
        }
    };
}
