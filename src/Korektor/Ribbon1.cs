using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;

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
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_LAT)
            {
                Properties.Settings.Default.Alphabet = "Latin";
            }
            else 
            {
                Properties.Settings.Default.Alphabet = "Cyrillic";
            }
        }

        private void correctErrors()
        {
            Globals.ThisAddIn.Application.Selection.Collapse();
            initializeCorrectorForm();
            Globals.ThisAddIn.frmCorrectWord.StartPosition = FormStartPosition.CenterScreen;
            WordFinder wordFinder = new WordFinder();
            if (wordFinder.next())
            {
                Globals.ThisAddIn.frmCorrectWord.setWordFinder(wordFinder);
                Globals.ThisAddIn.frmCorrectWord.ShowDialog();
            }
            if (wordFinder.getCheckFullyCompleted())
            {
                wordFinder.handleCheckFullyCompleted();
            }
            Globals.ThisAddIn.Application.ActiveWindow.SetFocus();
            Globals.ThisAddIn.Application.ActiveWindow.Activate();
        }

        private void initializeCorrectorForm()
        {
            if (null == Globals.ThisAddIn.frmCorrectWord)
            {
                Globals.ThisAddIn.frmCorrectWord = new FormCorrectWord();
            }
            Globals.ThisAddIn.frmCorrectWord = new FormCorrectWord();
        }

        private void ribbonButtonCyr_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CYR;
            correctErrors();
        }

        private void ribbonButtonLat_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_LAT;
            correctErrors();
        }

        private void ribbonButtonCyrLat_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CYRLAT;
            correctErrors();
        }

    }

}