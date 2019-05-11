using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Korektor
{
    public class HunspellFacade
    {
        public HunspellFacade() {
        }

        public bool Spell(String str)
        {
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CYR)
            {
                return Globals.ThisAddIn.hunspellCir.Spell(str);
            }
            else if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_LAT)
            {
                return Globals.ThisAddIn.hunspellLat.Spell(str);
            }
            else {
                return
                    Globals.ThisAddIn.hunspellCir.Spell(str) ||
                    Globals.ThisAddIn.hunspellLat.Spell(str);
            }

        }

        public List<string> Suggest(String str)
        {
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CYR)
            {
                return Globals.ThisAddIn.hunspellCir.Suggest(str);
            }
            else if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_LAT)
            {
                return Globals.ThisAddIn.hunspellLat.Suggest(str);
            }
            else
            {
                Regex regex = new Regex(@"\p{IsCyrillic}+");
                Match match = regex.Match(str);
                if (match.Success)
                {
                    return Globals.ThisAddIn.hunspellCir.Suggest(str);
                }
                else {
                    return Globals.ThisAddIn.hunspellLat.Suggest(str);
                }
            }
        }

        public void Add(String str)
        {
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CYR)
            {
                Globals.ThisAddIn.hunspellCir.Add(str);
            }
            else if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_LAT)
            {
                Globals.ThisAddIn.hunspellCir.Add(str);
            }
            else
            {
                Regex regex = new Regex(@"\p{IsCyrillic}+");
                Match match = regex.Match(str);
                if (match.Success)
                {
                    Globals.ThisAddIn.hunspellCir.Add(str);
                }
                else
                {
                    Globals.ThisAddIn.hunspellLat.Add(str);
                }
            }
        }

    }

}
