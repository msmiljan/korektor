using System;
using Microsoft.Office.Tools.Ribbon;

namespace Korektor
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            string alphabet = Properties.Settings.Default.Alphabet;
            if (alphabet.Equals("Cyrillic"))
            {
                Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CYR;
            }
            else
            {
                Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_LAT;
            }
        }

        private void Ribbon1_Close(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CYR)
            {
                Properties.Settings.Default.Alphabet = "Cyrillic";
            }
            else 
            {
                Properties.Settings.Default.Alphabet = "Latin";
            }
        }

        private void hyphenate()
        {
            Globals.ThisAddIn.hyphenator.hyphenate();
        }

        private void unhyphenate()
        {
            Globals.ThisAddIn.hyphenator.unhyphenate();
        }

        private void splitButtonCyr_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CYR;
            Globals.ThisAddIn.correctErrors();
        }

        private void buttonCyrToLat_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.convertor.convert(ThisAddIn.ALPHABET_LAT);
        }

        private void buttonCyrPodeliReci_click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CYR;
            hyphenate();
        }

        private void buttonCyrSastaviReci_click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CYR;
            unhyphenate();
        }

        private void splitButtonLat_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_LAT;
            Globals.ThisAddIn.correctErrors();
        }

        private void buttonLatToCyr_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.convertor.convert(ThisAddIn.ALPHABET_CYR);
        }

        private void buttonLatPodeliReci_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_LAT;
            hyphenate();
        }

        private void buttonLatSastaviReci_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_LAT;
            unhyphenate();
        }

        private void ribbonButtonLatToLat_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CUT_LAT;
            Globals.ThisAddIn.convertor.convertCutLatToLat();
        }

    }

}