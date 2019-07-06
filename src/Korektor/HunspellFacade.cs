using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Korektor
{
    public class HunspellFacade
    {
        public HunspellFacade() {
        }

        public bool Spell(string str)
        {
            str = str.Replace("\u001f", "");
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CYR)
            {
                return Globals.ThisAddIn.hunspellCir.Spell(str);
            }
            else if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_LAT)
            {
                return Globals.ThisAddIn.hunspellLat.Spell(str);
            }
            else if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CUT_LAT)
            {
                return Globals.ThisAddIn.convertor.spellCutLat(str);
            }
            return true;
        }

        public List<string> Suggest(string str)
        {
            str = str.Replace("\u001f", "");
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CYR)
            {
                return Globals.ThisAddIn.hunspellCir.Suggest(str);
            }
            else if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_LAT)
            {
                return Globals.ThisAddIn.hunspellLat.Suggest(str);
            }
            else if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CUT_LAT)
            {
                return Globals.ThisAddIn.convertor.suggestCutLat(str);
            }
            return null;
        }

        public void Add(string str)
        {
            str = str.Replace("\u001f", "");
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CYR)
            {
                Globals.ThisAddIn.hunspellCir.Add(str);
            }
            else
            {
                Globals.ThisAddIn.hunspellLat.Add(str);
            }
            //else
            //{
            //    Regex regex = new Regex(@"\p{IsCyrillic}+");
            //    Match match = regex.Match(str);
            //    if (match.Success)
            //    {
            //        Globals.ThisAddIn.hunspellCir.Add(str);
            //    }
            //    else
            //    {
            //        Globals.ThisAddIn.hunspellLat.Add(str);
            //    }
            //}
        }

    }

}
