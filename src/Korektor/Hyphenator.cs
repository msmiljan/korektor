using Microsoft.Office.Interop.Word;
using System;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Korektor
{
    public class Hyphenator
    {
        public void hyphenate()
        {
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                Selection selection = Globals.ThisAddIn.Application.Selection;
                if (selection.Start == selection.End)
                {
                    Globals.ThisAddIn.Application.ActiveWindow.Selection.WholeStory();
                    selection = Globals.ThisAddIn.Application.Selection;
                }
                else while (selection.Text.EndsWith("\r") || selection.Text.EndsWith("\n")) selection.End--;

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(selection.Range.WordOpenXML);
                XmlNamespaceManager xmlNamespaceManager = new XmlNamespaceManager(xmlDocument.NameTable);
                xmlNamespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                string undoTitle = "Подела на слогове.";
                if (Globals.ThisAddIn.alphabet != ThisAddIn.ALPHABET_CYR)
                {
                    undoTitle = "Podelа na slogove";
                }
                object objUndo = null;
                if (Globals.ThisAddIn.getOfficeVersion() > 12)
                {
                    objUndo = Globals.ThisAddIn.Application.UndoRecord;
                    ((UndoRecord)objUndo).StartCustomRecord(undoTitle);
                }

                try
                {
                    foreach (XmlNode xmlNode in xmlDocument.SelectNodes("//w:t", xmlNamespaceManager))
                    {
                        string replacedEqualSign = xmlNode.InnerText.Replace("=", "\u0001");
                        xmlNode.InnerText = hypenateText(replacedEqualSign);
                    }
                    object value = Missing.Value;
                    selection.InsertXML(xmlDocument.InnerXml, ref value);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("   " + ex.Message);

                }

                Globals.ThisAddIn.Application.ActiveWindow.Selection.Select();
                selection = Globals.ThisAddIn.Application.Selection;

                searchAndReplaceAll("=", "\u001f");
                searchAndReplaceAll("\u0001", "=");


                selection.End = selection.Start;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.ScreenRefresh();

                if (Globals.ThisAddIn.getOfficeVersion() > 12)
                {
                    ((UndoRecord)objUndo).EndCustomRecord();
                }

                string checkCompletedMessage = "Подела на слогове завршена.";
                if (Globals.ThisAddIn.alphabet != ThisAddIn.ALPHABET_CYR)
                {
                    checkCompletedMessage = "Podela na slogove završena.";
                }
                System.Windows.Forms.MessageBox.Show("   " + checkCompletedMessage);
            }
            catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show(ex.Message + " - " + ex.StackTrace);
            }
            finally
            {
                Selection selection = Globals.ThisAddIn.Application.Selection;
                selection.End = selection.Start;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.ScreenRefresh();
            }
        }

        public string hypenateText(string text)
        {
            string pattern = @"(?i)([абвгдђежзијклљмнњопрстћуфхцчџш]+)";
            if (Globals.ThisAddIn.alphabet != ThisAddIn.ALPHABET_CYR)
            {
                pattern = @"(?i)([abcčćdđefghijklmnoprsštuvzž]+)";
            }
            Regex rgx = new Regex(pattern);

            StringBuilder sb = new StringBuilder(text.Length);
            foreach (string result in Regex.Split(text, pattern))
            {
                if (string.IsNullOrEmpty(result))
                {
                    continue;
                }
                sb.Append(Globals.ThisAddIn.getHyphen().Hyphenate(result).HyphenatedWord);
            }

            return sb.ToString();
        }

        public void unhyphenate()
        {
            try
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                searchAndReplaceAll("\u001f", "");

                Selection selection = Globals.ThisAddIn.Application.Selection;
                selection.End = selection.Start;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.ScreenRefresh();

                string checkCompletedMessage = "Састављање слогова завршено.";
                if (Globals.ThisAddIn.alphabet != ThisAddIn.ALPHABET_CYR)
                {
                    checkCompletedMessage = "Sastavljanje slogova završeno.";
                }
                System.Windows.Forms.MessageBox.Show("   " + checkCompletedMessage);
            }
            finally
            {
                Selection selection = Globals.ThisAddIn.Application.Selection;
                selection.End = selection.Start;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.ScreenRefresh();
            }
        }

        public static void searchAndReplaceAll(string from, string to)
        {
            object missing = System.Type.Missing;

            try
            {
                Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                // Loop through the StoryRanges (sections of the Word doc)
                foreach (Range tmpRange in doc.StoryRanges)
                {

                    // Set the text to find and replace
                    string text = tmpRange.Text;

                    object replaceAll = WdReplace.wdReplaceAll;
                    object matchCase = false;
                    object matchWholeWord = false;
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
                    object replace = WdReplace.wdReplaceAll;
                    object wrap = WdFindWrap.wdFindContinue;
                    //execute find and replace
                    tmpRange.Find.Execute(from, ref matchCase, ref matchWholeWord,
                        ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, to, ref replace,
                        ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                }
                foreach (ContentControl control in doc.ContentControls)
                {
                    // if (tmpControl.LockContentControl || tmpControl.LockContents || tmpControl.Temporary) continue;
                    control.Range.Text = control.Range.Text.Replace(from, to);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("   " + ex.Message);
                System.Windows.Forms.MessageBox.Show("   " + ex.StackTrace);
            }
        }

    }
}
