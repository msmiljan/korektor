using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Xml;
using System;
using System.Text.RegularExpressions;

namespace Korektor
{
    public class Convertor
    {
        private static Dictionary<string, string> mapCyrToLat = new Dictionary<string, string>();
        private static Dictionary<string, string> mapLatToCyr = new Dictionary<string, string>();

        static Convertor()
        {
            mapCyrToLat.Add("а", "a");
            mapCyrToLat.Add("б", "b");
            mapCyrToLat.Add("в", "v");
            mapCyrToLat.Add("г", "g");
            mapCyrToLat.Add("д", "d");
            mapCyrToLat.Add("ђ", "đ");
            mapCyrToLat.Add("е", "e");
            mapCyrToLat.Add("ж", "ž");
            mapCyrToLat.Add("з", "z");
            mapCyrToLat.Add("и", "i");
            mapCyrToLat.Add("ј", "j");
            mapCyrToLat.Add("к", "k");
            mapCyrToLat.Add("л", "l");
            mapCyrToLat.Add("љ", "lj");
            mapCyrToLat.Add("м", "m");
            mapCyrToLat.Add("н", "n");
            mapCyrToLat.Add("њ", "nj");
            mapCyrToLat.Add("о", "o");
            mapCyrToLat.Add("п", "p");
            mapCyrToLat.Add("р", "r");
            mapCyrToLat.Add("с", "s");
            mapCyrToLat.Add("т", "t");
            mapCyrToLat.Add("ћ", "ć");
            mapCyrToLat.Add("у", "u");
            mapCyrToLat.Add("ф", "f");
            mapCyrToLat.Add("х", "h");
            mapCyrToLat.Add("ц", "c");
            mapCyrToLat.Add("ч", "č");
            mapCyrToLat.Add("џ", "dž");
            mapCyrToLat.Add("ш", "š");

            mapCyrToLat.Add("А", "A");
            mapCyrToLat.Add("Б", "B");
            mapCyrToLat.Add("В", "V");
            mapCyrToLat.Add("Г", "G");
            mapCyrToLat.Add("Д", "D");
            mapCyrToLat.Add("Ђ", "Đ");
            mapCyrToLat.Add("Е", "E");
            mapCyrToLat.Add("Ж", "Ž");
            mapCyrToLat.Add("З", "Z");
            mapCyrToLat.Add("И", "I");
            mapCyrToLat.Add("Ј", "J");
            mapCyrToLat.Add("К", "K");
            mapCyrToLat.Add("Л", "L");
            mapCyrToLat.Add("Љ", "Lj");
            mapCyrToLat.Add("М", "M");
            mapCyrToLat.Add("Н", "N");
            mapCyrToLat.Add("Њ", "Nj");
            mapCyrToLat.Add("О", "O");
            mapCyrToLat.Add("П", "P");
            mapCyrToLat.Add("Р", "R");
            mapCyrToLat.Add("С", "S");
            mapCyrToLat.Add("Т", "T");
            mapCyrToLat.Add("Ћ", "Ć");
            mapCyrToLat.Add("У", "U");
            mapCyrToLat.Add("Ф", "F");
            mapCyrToLat.Add("Х", "H");
            mapCyrToLat.Add("Ц", "C");
            mapCyrToLat.Add("Ч", "Č");
            mapCyrToLat.Add("Џ", "Dž");
            mapCyrToLat.Add("Ш", "Š");

            mapLatToCyr.Add("a", "а");
            mapLatToCyr.Add("b", "б");
            mapLatToCyr.Add("v", "в");
            mapLatToCyr.Add("g", "г");
            mapLatToCyr.Add("d", "д");
            mapLatToCyr.Add("đ", "ђ");
            mapLatToCyr.Add("e", "е");
            mapLatToCyr.Add("ž", "ж");
            mapLatToCyr.Add("z", "з");
            mapLatToCyr.Add("i", "и");
            mapLatToCyr.Add("j", "ј");
            mapLatToCyr.Add("k", "к");
            mapLatToCyr.Add("l", "л");
            mapLatToCyr.Add("lj", "љ");
            mapLatToCyr.Add("m", "м");
            mapLatToCyr.Add("n", "н");
            mapLatToCyr.Add("nj", "њ");
            mapLatToCyr.Add("o", "о");
            mapLatToCyr.Add("p", "п");
            mapLatToCyr.Add("r", "р");
            mapLatToCyr.Add("s", "с");
            mapLatToCyr.Add("t", "т");
            mapLatToCyr.Add("ć", "ћ");
            mapLatToCyr.Add("u", "у");
            mapLatToCyr.Add("f", "ф");
            mapLatToCyr.Add("h", "х");
            mapLatToCyr.Add("c", "ц");
            mapLatToCyr.Add("č", "ч");
            mapLatToCyr.Add("dž", "џ");
            mapLatToCyr.Add("š", "ш");

            mapLatToCyr.Add("A", "А");
            mapLatToCyr.Add("B", "Б");
            mapLatToCyr.Add("V", "В");
            mapLatToCyr.Add("G", "Г");
            mapLatToCyr.Add("D", "Д");
            mapLatToCyr.Add("Đ", "Ђ");
            mapLatToCyr.Add("E", "Е");
            mapLatToCyr.Add("Ž", "Ж");
            mapLatToCyr.Add("Z", "З");
            mapLatToCyr.Add("I", "И");
            mapLatToCyr.Add("J", "Ј");
            mapLatToCyr.Add("K", "К");
            mapLatToCyr.Add("L", "Л");
            mapLatToCyr.Add("Lj", "Љ");
            mapLatToCyr.Add("M", "М");
            mapLatToCyr.Add("N", "Н");
            mapLatToCyr.Add("Nj", "Њ");
            mapLatToCyr.Add("O", "О");
            mapLatToCyr.Add("P", "П");
            mapLatToCyr.Add("R", "Р");
            mapLatToCyr.Add("S", "С");
            mapLatToCyr.Add("T", "Т");
            mapLatToCyr.Add("Ć", "Ћ");
            mapLatToCyr.Add("U", "У");
            mapLatToCyr.Add("F", "Ф");
            mapLatToCyr.Add("H", "Х");
            mapLatToCyr.Add("C", "Ц");
            mapLatToCyr.Add("Č", "Ч");
            mapLatToCyr.Add("Dž", "Џ");
            mapLatToCyr.Add("Š", "Ш");
        }

        private string toCyrillic(string text)
        {
            int i;
            char chr;
            StringBuilder sb = new StringBuilder();
            for (i = 0; i < text.Length; i++)
            {
                string textLower = text.ToLower();
                if (i < text.Length - 1)
                {
                    if (!(textLower[i] == 'n' || textLower[i] == 'l' ? textLower[i + 1] != 'j' : true))
                    {
                        sb.Append(mapLatToCyr[string.Concat(text[i], text[i + 1])]);
                        i++;
                        continue;
                    }
                    else if (!(textLower[i] == 'd' ? textLower[i + 1] != 'ž' : true))
                    {
                        sb.Append(mapLatToCyr[string.Concat(text[i], text[i + 1])]);
                        i++;
                        continue;
                    }
                }
                if (!mapLatToCyr.ContainsKey(text[i].ToString()))
                {
                    chr = text[i];
                    sb.Append(chr.ToString());
                }
                else
                {
                    Dictionary<string, string> strs = mapLatToCyr;
                    chr = text[i];
                    sb.Append(strs[chr.ToString()]);
                }
            }
            return sb.ToString();
        }

        private string toLatin(string text)
        {
            string str;
            StringBuilder stringBuilder = new StringBuilder();
            string str1 = text;
            for (int i = 0; i < str1.Length; i++)
            {
                char chr = str1[i];
                str = (!mapCyrToLat.ContainsKey(chr.ToString()) ? chr.ToString() : mapCyrToLat[chr.ToString()]);
                stringBuilder.Append(str);
            }
            return stringBuilder.ToString();
        }

        public void convert(int alphabet)
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

                //Hyphenator.searchAndReplaceAll("\u001f", "");

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(selection.Range.WordOpenXML);
                XmlNamespaceManager xmlNamespaceManager = new XmlNamespaceManager(xmlDocument.NameTable);
                xmlNamespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                string undoTitle = "Конверзија писма.";
                if (alphabet != ThisAddIn.ALPHABET_CYR)
                {
                    undoTitle = "Konverzija pisma.";
                }
                object objUndo = null;
                if (Globals.ThisAddIn.getOfficeVersion() > 12) {
                    objUndo = Globals.ThisAddIn.Application.UndoRecord;
                    ((UndoRecord)objUndo).StartCustomRecord(undoTitle);
                }

                foreach (XmlNode xmlNode in xmlDocument.SelectNodes("//w:t", xmlNamespaceManager))
                {
                    xmlNode.InnerText = convertText(xmlNode.InnerText, alphabet);
                }
                selection.InsertXML(xmlDocument.InnerXml, Missing.Value);

                if (Globals.ThisAddIn.getOfficeVersion() > 12)
                {
                    ((UndoRecord)objUndo).EndCustomRecord();
                }

                string checkCompletedMessage = "Конверзија завршена.";
                if (alphabet != ThisAddIn.ALPHABET_CYR)
                {
                    checkCompletedMessage = "Konverzija završena.";
                }

                selection.End = selection.Start;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.ScreenRefresh();
                System.Windows.Forms.MessageBox.Show("   " + checkCompletedMessage);
            } 
            catch (Exception ex)
            {
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

        private string convertText(string text, int alphabet)
        {
            string convertedText = "";
            if (alphabet == ThisAddIn.ALPHABET_LAT)
            {
                convertedText = toLatin(text);
            }
            else if (alphabet == ThisAddIn.ALPHABET_CYR)
            {
                convertedText = toCyrillic(text);
                //convertedText = convertLatToCyrNJDZWithSpellCheck(convertedText);
            }
            return convertedText;
        }

        private string convertLatToCyrNJDZWithSpellCheck(string text)
        {
            string pattern = @"(?i)([абвгдђежзијклљмнњопрстћуфхцчџшqwxy]+)";
            Regex rgx = new Regex(pattern);
            StringBuilder sb = new StringBuilder(text.Length);
            foreach (string result in Regex.Split(text, pattern))
            {
                if (string.IsNullOrEmpty(result))
                {
                    continue;
                }
                if (result.ToLower().Contains("њ") && !Globals.ThisAddIn.getHunspell().Spell(result))
                {
                    sb.Append(result.Replace("њ", "нј"));
                }
                else if (result.ToLower().Contains("џ") && !Globals.ThisAddIn.getHunspell().Spell(result))
                {
                    sb.Append(result.Replace("џ", "дж"));
                }
                else
                {
                    sb.Append(result);
                }
            }
            return sb.ToString();
        }

        private HashSet<string> hairyWords = new HashSet<string>();

        public void convertCutLatToLat()
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

                Hyphenator.searchAndReplaceAll("\u001f", "");

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(selection.Range.WordOpenXML);
                XmlNamespaceManager xmlNamespaceManager = new XmlNamespaceManager(xmlDocument.NameTable);
                xmlNamespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                object objUndo = null;
                if (Globals.ThisAddIn.getOfficeVersion() > 12)
                {
                    objUndo = Globals.ThisAddIn.Application.UndoRecord;
                    ((UndoRecord)objUndo).StartCustomRecord("Konverzija pisma");
                }

                foreach (XmlNode xmlNode in xmlDocument.SelectNodes("//w:t", xmlNamespaceManager))
                {
                    xmlNode.InnerText = convertCutLatText(xmlNode.InnerText);
                }
                selection.InsertXML(xmlDocument.InnerXml, Missing.Value);

                if (Globals.ThisAddIn.getOfficeVersion() > 12)
                {
                    ((UndoRecord)objUndo).EndCustomRecord();
                }

                string checkCompletedMessage = "Automatski deo konverzije ošišane latinice završen.\r\n\r\nMolim proverite reči koje je moguće konvertovati na više načina...";
                selection.End = selection.Start;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.ScreenRefresh();

                if (selection.Text.Trim().Length > 0)
                {
                    System.Windows.Forms.MessageBox.Show("   " + checkCompletedMessage);
                }

                Globals.ThisAddIn.correctErrors();

            }
            catch (Exception ex)
            {
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

        private string convertCutLatText(string text)
        {
            text = convertCutLatDJWithSpellCheck(text);
            string pattern = @"(?i)([abcčćdđefghijklmnoprsštuvzž]+)";
            Regex rgx = new Regex(pattern);
            StringBuilder sb = new StringBuilder(text.Length);
            foreach (string result in Regex.Split(text, pattern))
            {
                if (string.IsNullOrEmpty(result))
                {
                    continue;
                }
                hairyWords.Clear();
                getHairyWords(result, 0, 0);
                if (hairyWords.Count == 1)
                {
                    sb.Append(result);
                }
                else
                {
                    List<string> correct = new List<string>();
                    Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_LAT;
                    foreach (string str in hairyWords)
                    {
                        if (Globals.ThisAddIn.getHunspell().Spell(str))
                        {
                            correct.Add(str);
                        }
                        if (correct.Count > 1) break; 
                    }
                    Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CUT_LAT;
                    if (correct.Count == 1)
                    {
                        sb.Append(correct[0]);
                    }
                    else
                    {
                        sb.Append(result);
                    }
                }
            }
            return sb.ToString();
        }

        private string convertCutLatDJWithSpellCheck(string text)
        {
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_LAT;

            string pattern = @"(?i)([abcčćdđefghijklmnoprsštuvzž]+)";
            Regex rgx = new Regex(pattern);
            StringBuilder sb = new StringBuilder(text.Length);
            foreach (string result in Regex.Split(text, pattern))
            {
                if (string.IsNullOrEmpty(result))
                {
                    continue;
                }
                if (result.ToLower().Contains("dj") && !Globals.ThisAddIn.getHunspell().Spell(result))
                {
                    sb.Append(result.Replace("dj", "đ").Replace("DJ", "Đ").Replace("Dj", "Đ"));
                }
                else
                {
                    sb.Append(result);
                }
            }
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CUT_LAT;
            return sb.ToString();
        }

        private string replaceCharAt(string str, int index, char ch)
        {
            char[] chars = str.ToCharArray();
            chars[index] = ch;
            return new string(chars);
        }

        private void getHairyWords(string word, int index1, int index2)
        {
            if (index1 > word.Length - 1)
            {
                return;
            }
            char c2 = word[index2];

            switch (c2)
            {
                case 'z':
                    hairyWords.Add(word);
                    getHairyWords(word, index1 + 1, index2 + 1);
                    word = replaceCharAt(word, index2, 'ž');
                    getHairyWords(word, index1, index2);
                    break;
                case 'Z':
                    hairyWords.Add(word);
                    getHairyWords(word, index1 + 1, index2 + 1);
                    word = replaceCharAt(word, index2, 'Ž');
                    getHairyWords(word, index1, index2);
                    break;
                case 's':
                    hairyWords.Add(word);
                    getHairyWords(word, index1 + 1, index2 + 1);
                    word = replaceCharAt(word, index2, 'š');
                    getHairyWords(word, index1, index2);
                    break;
                case 'S':
                    hairyWords.Add(word);
                    getHairyWords(word, index1 + 1, index2 + 1);
                    word = replaceCharAt(word, index2, 'Š');
                    getHairyWords(word, index1, index2);
                    break;
                case 'c':
                    hairyWords.Add(word);
                    getHairyWords(word, index1 + 1, index2 + 1);
                    word = replaceCharAt(word, index2, 'ć');
                    getHairyWords(word, index1, index2);
                    word = replaceCharAt(word, index2, 'č');
                    getHairyWords(word, index1, index2);
                    break;
                case 'C':
                    hairyWords.Add(word);
                    getHairyWords(word, index1 + 1, index2 + 1);
                    word = replaceCharAt(word, index2, 'Ć');
                    getHairyWords(word, index1, index2);
                    word = replaceCharAt(word, index2, 'Č');
                    getHairyWords(word, index1, index2);
                    break;
                case 'ž':
                case 'Ž':
                case 'š':
                case 'Š':
                case 'ć':
                case 'č':
                case 'Ć':
                case 'Č':
                    hairyWords.Add(word);
                    getHairyWords(word, index1 + 1, index2 + 1);
                    break;
                default:
                    hairyWords.Add(word);
                    getHairyWords(word, index1 + 1, index2 + 1);
                    break;
            }

        }

        public bool spellCutLat(string word)
        {
            hairyWords.Clear();
            getHairyWords(word, 0, 0);
            if (hairyWords.Count < 2)
            {
                return true;
            }
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_LAT;
            int i = 0;
            foreach (string str in hairyWords)
            {
                if (Globals.ThisAddIn.getHunspell().Spell(str))
                {
                    if (i++ > 1) break;
                }
            }
            Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CUT_LAT;
            return i < 2;
        }

        public List<string> suggestCutLat(string word)
        {
            try
            {
                hairyWords.Clear();
                getHairyWords(word, 0, 0);
                List<string> correct = new List<string>();
                Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_LAT;
                foreach (string str in hairyWords)
                {
                    if (Globals.ThisAddIn.getHunspell().Spell(str))
                    {
                        correct.Add(str);
                    }
                }
                return correct;
            }
            finally
            {
                Globals.ThisAddIn.alphabet = ThisAddIn.ALPHABET_CUT_LAT;
            }
        }

    }
}
