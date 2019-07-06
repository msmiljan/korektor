using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Korektor
{
    public class WordFinder
    {

        // list of story ranges to check
        List<Word.Range> storyRanges = new List<Word.Range>();

        // current story range index
        int currentStoryRangeIndex = 0;

        // current story range
        Word.Range currentStoryRange = null;

        // initially from cursor to document end
        Word.Range rngBottom = null;

        // initially from document start to cursor
        Word.Range rngTop = null;

        // time to switch between bottom and top correction range?
        bool endPassed = false;

        // last checked word length
        int lastWordLength = 0;

        // last checked word end
        int lastWordEndPos = 0;

        // current incorrect word
        string currentWord = null;

        // story check fully completed
        private bool storyCheckCompleted = false;

        // check fully completed
        private bool checkFullyCompleted = false;

        public WordFinder()
        {
            this.initializeCorrectionRanges();
        }

        public bool next()
        {
            this.currentWord = null;
            checkIncorrectSpelling();
            return !string.IsNullOrEmpty(this.getCurrentWord());
        }

        public string getCurrentWord()
        {
            return this.currentWord;
        }

        private void setCurrentWord(string currentWord)
        {
            this.currentWord = currentWord;
        }

        public bool getStoryCheckCompleted()
        {
            return this.storyCheckCompleted;
        }

        public void setStoryCheckCompleted(bool storyCheckCompleted)
        {
            this.storyCheckCompleted = storyCheckCompleted;
        }

        public bool getCheckFullyCompleted()
        {
            return this.checkFullyCompleted;
        }

        public void setCheckFullyCompleted(bool checkFullyCompleted)
        {
            this.checkFullyCompleted = checkFullyCompleted;
        }

        /// finds incorrect if exists, and sets it as current word
        public bool checkIncorrectSpelling()
        {
            Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorWait;
            try
            {
                bool wordFound = false;
                do
                {
                    Word.Range rng = getCorrectionRange();
                    if (null == rng)
                    {
                        setCheckFullyCompleted(true);
                        return false;
                    }
                    foreach (Word.Range w in rng.Words)
                    {
                        if (null != w.ContentControls && w.ContentControls.Count > 0)
                        {
                            continue;
                        }
                        if (w.End > rng.End)
                        {
                            break;
                        }
                        this.lastWordLength = w.End - w.Start;
                        this.lastWordEndPos = w.End;
                        string str = w.Text.Trim();
                        while (!string.IsNullOrEmpty(str) && !Char.IsLetter(str[str.Length - 1]))
                        {
                            str = str.TrimEnd(str[str.Length - 1]);
                        }
                        if (!string.IsNullOrEmpty(str) &&
                            !Globals.ThisAddIn.ignoreAllWords.Contains(str) &&
                            !Globals.ThisAddIn.getHunspell().Spell(str))
                        {
                            setCurrentWord(str);
                            w.SetRange(w.Start, w.Start + str.Length);
                            w.Select();
                            wordFound = true;
                            break;
                        }
                    }
                    if (!wordFound)
                    {
                        if (this.endPassed == true)
                        {
                            this.setStoryCheckCompleted(true);
                            this.endPassed = false;
                        }
                        else
                        {
                            this.endPassed = true;
                        }
                    }
                    else
                    {
                        this.advanceCorrectionRange();
                        break;
                    }
                } while (true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace);
            }
            finally
            {
                Globals.ThisAddIn.Application.System.Cursor = Word.WdCursorType.wdCursorNormal;
                Globals.ThisAddIn.Application.ScreenRefresh();
            }
            return true;
        }

        /// displays completed dialog
        /// collapses selection
        public void handleCheckFullyCompleted()
        {
            Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            string checkFullyCompletedMessage = null;
            if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CYR)
            {
                checkFullyCompletedMessage = "Провера завршена.";
            }
            else if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_LAT)
            {
                checkFullyCompletedMessage = "Provera završena.";
            }
            else if (Globals.ThisAddIn.alphabet == ThisAddIn.ALPHABET_CUT_LAT)
            {
                checkFullyCompletedMessage = "Konverzija završena.";
            }
            MessageBox.Show("   " + checkFullyCompletedMessage);
        }

        /// handles change all event
        public static void findAndReplace(Microsoft.Office.Interop.Word.Selection selection, object findText, object replaceWithText)
        {
            // TODO include all editable stories
            //options
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        ///  add word to user dictionary
        public void addToDictionary(string str)
        {
            Globals.ThisAddIn.addToCustomDictionaryFile(str);
            Globals.ThisAddIn.getHunspell().Add(str);
        }

        ///  get text for current selection
        private string getSelectionText()
        {
            return Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.Selection.Text;
        }

        // finds all storyRanges and sets those for checking
        // curent story range is excluded from list and made starting one
        public void initializeCorrectionRanges()
        {
            storyRanges.Clear();
            this.endPassed = false;
            Word.Application app = Globals.ThisAddIn.Application;
            foreach (Word.Range stoRng in Globals.ThisAddIn.Application.ActiveDocument.StoryRanges)
            {
                if ( isStoryTypeHeaderOrFooter(stoRng) ||
                     isStoryTypeNoteSeparator(stoRng) )
                {
                    continue;
                }
                Word.Range stoRngCurrent = stoRng;
                if (app.Selection.InStory(stoRngCurrent))
                {
                    currentStoryRange = stoRngCurrent;
                }
                else
                {
                    storyRanges.Add(stoRngCurrent);
                }
                do
                {
                    stoRngCurrent = stoRngCurrent.NextStoryRange;
                    if (null != stoRngCurrent)
                    {
                        if (app.Selection.InStory(stoRngCurrent))
                        {
                            this.currentStoryRange = stoRngCurrent;
                        }
                        else
                        {
                            storyRanges.Add(stoRngCurrent);
                        }
                    }
                } while (null != stoRngCurrent);
            }
            this.rngTop = app.Selection.Range;
            this.rngTop.SetRange(this.currentStoryRange.Start, app.Selection.Range.Start);
            this.rngBottom = app.Selection.Range;
            this.rngBottom.SetRange(app.Selection.Range.End, this.currentStoryRange.End);
        }

        /// recalculates correction range after each word found/handled
        /// moves calculation to next story range if current one is completed
        private Word.Range getCorrectionRange()
        {
            if (this.getStoryCheckCompleted()) {
                if (this.currentStoryRangeIndex < this.storyRanges.Count)
                {
                    Word.Range stoRngCurrent = this.storyRanges[this.currentStoryRangeIndex];
                    this.currentStoryRange = stoRngCurrent;
                    this.currentStoryRangeIndex++;
                    Word.Application app = Globals.ThisAddIn.Application;
                    Word.Range rngSelect = stoRngCurrent.Duplicate;
                    rngSelect.SetRange(rngSelect.Start, rngSelect.Start);
                    rngSelect.Select();
                    this.rngTop = stoRngCurrent.Duplicate;
                    this.rngTop.SetRange(stoRngCurrent.Start, stoRngCurrent.End);
                    this.rngBottom = stoRngCurrent.Duplicate;
                    this.rngBottom.SetRange(rngBottom.End, rngBottom.End);
                    this.setStoryCheckCompleted(false);
                }
                else
                {
                    this.rngTop = null;
                    this.rngBottom = null;
                }
            }
            Word.Range rng = this.rngBottom;
            if (this.endPassed)
            {
                rng = this.rngTop;
            }
            return rng;
        }

        /// as new incorrect word is found, move correction range start behind that word
        private void advanceCorrectionRange()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            if (!this.endPassed)
            {
                this.rngBottom.SetRange(this.lastWordEndPos, this.currentStoryRange.End);
            }
            else
            {
                this.rngTop.SetRange(this.lastWordEndPos, this.rngTop.End);
            }
        }


        /// as new incorrect word is replaced with correct word, adjust correction range for old/new word length diff
        public void adjustCorrectionRangeOnChange()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app.ActiveDocument;
            if (!this.endPassed)
            {
                this.rngBottom.SetRange(this.rngBottom.Start + this.lastWordLength - app.Selection.Text.Length, doc.Content.End);
            }
            else
            {
                this.rngTop.SetRange(this.lastWordEndPos + 1, this.rngTop.End);
            }
        }

        /// is Story Type Header Or Footer
        private bool isStoryTypeHeaderOrFooter(Word.Range range)
        {
            return (range.StoryType == Word.WdStoryType.wdEvenPagesHeaderStory ||
                    range.StoryType == Word.WdStoryType.wdPrimaryHeaderStory ||
                    range.StoryType == Word.WdStoryType.wdEvenPagesFooterStory ||
                    range.StoryType == Word.WdStoryType.wdPrimaryFooterStory ||
                    range.StoryType == Word.WdStoryType.wdFirstPageHeaderStory ||
                    range.StoryType == Word.WdStoryType.wdFirstPageFooterStory);
        }

        /// is story type note separator
        private bool isStoryTypeNoteSeparator(Word.Range range)
        {
            return (range.StoryType == Word.WdStoryType.wdFootnoteSeparatorStory ||
                    range.StoryType == Word.WdStoryType.wdFootnoteContinuationSeparatorStory ||
                    range.StoryType == Word.WdStoryType.wdEndnoteSeparatorStory ||
                    range.StoryType == Word.WdStoryType.wdEndnoteContinuationSeparatorStory);
        }

    }

}
