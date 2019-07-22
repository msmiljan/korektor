using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Korektor
{
    public partial class FormCorrectWord : Form
    {

        private WordFinder wordFinder;

        public FormCorrectWord()
        {
            InitializeComponent();
        }

        public void setWordFinder(WordFinder wordFinder)
        {
            this.wordFinder = wordFinder;
        }

        private WordFinder getWordFinder()
        {
            return this.wordFinder;
        }

        private void btnIgnore_Click(object sender, EventArgs e)
        {
            //gotoNextWord();
            if (!checkNextWord())
            {
                checkCompleted();
            }
        }

        private void btnIgnoreAll_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.ignoreAllWords.Add(getSelectionText());
            //gotoNextWord();
            if (!checkNextWord())
            {
                checkCompleted();
            }
        }

        private void btnAddToDictionary_Click(object sender, EventArgs e)
        {
            addToDictionary(getSelectionText());
            //gotoNextWord();
            if (!checkNextWord())
            {
                checkCompleted();
            }
        }

        private void lstSuggestions_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstSuggestions.SelectedItem != null)
            {
                this.tbxReplace.Text = lstSuggestions.SelectedItem.ToString();
            }
        }

        private void tbxReplace_TextChanged(object sender, EventArgs e)
        {
            setControlsEnablement();
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.Selection.Text = this.tbxReplace.Text;
            this.getWordFinder().adjustCorrectionRangeOnChange();
            //gotoNextWord();
            if (!checkNextWord())
            {
                checkCompleted();
            }
        }

        private void btnChangeAll_Click(object sender, EventArgs e)
        {
            WordFinder.findAndReplace(Globals.ThisAddIn.Application.Selection,
                Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.Selection.Text,
                this.tbxReplace.Text);
            getWordFinder().initializeCorrectionRanges();
            if (!checkNextWord())
            {
                checkCompleted();
            }
        }

        private void FormCorrectWord_Load(object sender, EventArgs e)
        {
            this.setFormUIAlphabet();
            populateWordFields();
            if (Globals.ThisAddIn.frmCorrectWordLocation.X == 0 && Globals.ThisAddIn.frmCorrectWordLocation.Y == 0)
            {
                this.StartPosition = FormStartPosition.CenterScreen;
            }
            else
            {
                this.Location = Globals.ThisAddIn.frmCorrectWordLocation;
            }
        }

        private void FormCorrectWord_FormClosed(object sender, FormClosedEventArgs e)
        {
            Globals.ThisAddIn.frmCorrectWordLocation = this.Location;
            Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        private void FormCorrectWord_VisibleChanged(object sender, EventArgs e)
        {
        }

        private void FormCorrectWord_Activated(object sender, EventArgs e)
        {
        }

        public void checkCurrentWord()
        {
            populateWordFields();
        }

        public bool checkNextWord()
        {
            if (getWordFinder().next())
            {
                populateWordFields();
                return true;
            }
            return false;
        }

        private void setControlsEnablement()
        {
            if (this.tbxReplace.Text != getSelectionText())
            {
                this.btnChange.Enabled = true;
                this.btnChangeAll.Enabled = true;
            }
            else
            {
                this.btnChange.Enabled = false;
                this.btnChangeAll.Enabled = false;
            }
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CUT_LAT)
            {
                this.btnAddToDictionary.Enabled = false;
            }
            else
            {
                this.btnAddToDictionary.Enabled = true;
            }
        }

        private void addToDictionary(string str)
        {
            Globals.ThisAddIn.addToCustomDictionaryFile(str);
            Globals.ThisAddIn.getHunspell().Add(str);
        }

        private string getSelectionText()
        {
            return Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.Selection.Text;
        }

        private void setFormUIAlphabet()
        {
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CYR)
            {
                this.Text = "Реч";
                this.btnIgnore.Text = "Прескочи";
                this.btnIgnoreAll.Text = "Прескочи све";
                this.btnChange.Text = "Замени";
                this.btnChangeAll.Text = "Замени све";
                this.btnAddToDictionary.Text = "Додај у речник";
            }
            else 
            {
                this.Text = "Reč";
                this.btnIgnore.Text = "Preskoči";
                this.btnIgnoreAll.Text = "Preskoči sve";
                this.btnChange.Text = "Zameni";
                this.btnChangeAll.Text = "Zameni sve";
                this.btnAddToDictionary.Text = "Dodaj u rečnik";
            }
        }

        // populates textbox and suggestions for incorrect word 
        private void populateWordFields()
        {
            string word = getWordFinder().getCurrentWord();
            this.tbxReplace.Text = word;
            List<string> suggestions = Globals.ThisAddIn.getHunspell().Suggest(word);
            this.lstSuggestions.Items.Clear();
            for (int i = 0; i < suggestions.Count; i++)
            {
                this.lstSuggestions.Items.Add(suggestions.ElementAt(i));
            }
            setControlsEnablement();
        }

        private void checkCompleted()
        {
            getWordFinder().setCheckFullyCompleted(true);
            Close();
        }

    }

}
