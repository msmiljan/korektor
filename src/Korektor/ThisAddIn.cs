using NHunspell;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace Korektor
{
    public partial class ThisAddIn
    {

        public static string CUSTOM_DICT_PREFIX = "CustomWords_";
        public static string FILE_CYR = "sr";
        public static string FILE_LAT = "sr-Latn";
        public static string PROGRAM_FOLDER = "Korektor";

        public static int ALPHABET_CYR = 1;
        public static int ALPHABET_LAT = 2;
        public static int ALPHABET_CYRLAT = 3;

        public HunspellFacade hunspellFacade = null;

        public Hunspell hunspellCir = null;
        public Hunspell hunspellLat = null;

        public string programAppData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + PROGRAM_FOLDER;
        public int alphabet = 1;

        public FormCorrectWord frmCorrectWord = new FormCorrectWord();

        public HashSet<string> ignoreAllWords = new HashSet<string>();

        public string dictionariesPath = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            setDeploymentPath();
            initializeHunspell();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void initializeHunspell()
        {
            if (!Directory.Exists(this.programAppData))
            {
                Directory.CreateDirectory(this.programAppData);
            }
            this.hunspellCir = new Hunspell(this.dictionariesPath + FILE_CYR + ".aff", this.dictionariesPath + FILE_CYR + ".dic");
            string customPathCir = this.programAppData + "\\" + CUSTOM_DICT_PREFIX + FILE_CYR + ".txt";
            if (!File.Exists(customPathCir))
            {
                File.CreateText(customPathCir).Close();
            }
            string[] lines = System.IO.File.ReadAllLines(customPathCir);
            foreach (var line in lines)
            {
                hunspellCir.Add(line);
            }

            this.hunspellLat = new Hunspell(this.dictionariesPath + FILE_LAT + ".aff", this.dictionariesPath + FILE_LAT + ".dic");
            string customPathLat = this.programAppData + "\\" + CUSTOM_DICT_PREFIX + FILE_LAT + ".txt";
            if (!File.Exists(customPathLat))
            {
                File.CreateText(customPathLat).Close();
            }
            lines = System.IO.File.ReadAllLines(customPathLat);
            foreach (var line in lines)
            {
                hunspellLat.Add(line);
            }
            hunspellFacade = new HunspellFacade();
    }

    public void setDeploymentPath()
        {
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

            //Location is where the assembly is run from 
            string assemblyLocation = assemblyInfo.Location;

            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            this.dictionariesPath = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString()) + "\\";
        }

        public void addToCustomDictionaryFile(string str)
        {
            String file = "";
            if (this.alphabet == ALPHABET_LAT)
            {
                file = this.programAppData + "\\" + CUSTOM_DICT_PREFIX + FILE_LAT + ".txt"; ;
            }
            else if (this.alphabet == ALPHABET_CYR)
            {
                file = this.programAppData + "\\" + CUSTOM_DICT_PREFIX + FILE_CYR + ".txt";
            }
            else {
                Regex regex = new Regex(@"\p{IsCyrillic}+");
                Match match = regex.Match(str);
                if (match.Success)
                {
                    file = this.programAppData + "\\" + CUSTOM_DICT_PREFIX + FILE_CYR + ".txt";
                }
                else
                {
                    file = this.programAppData + "\\" + CUSTOM_DICT_PREFIX + FILE_LAT + ".txt"; ;
                }
            }
            using (StreamWriter sw = File.AppendText(file))
            {
                sw.WriteLine(str);
                sw.Close();
            }
        }

        public HunspellFacade getHunspell()
        {
            return hunspellFacade;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
