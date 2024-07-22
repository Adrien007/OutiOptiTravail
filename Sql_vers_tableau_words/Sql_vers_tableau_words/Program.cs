
/* 
MIT License

Copyright (c) 2024 Adrien Choiniere

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/




using OfficeOpenXml;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Media;
using System.Xml.Linq;

namespace Sql_vers_tableau_words
{
    internal class Program
    {
        public const string BALISE_TAB = "--%%%C_TabBalise";
        public const string RESULTAT_ATTENDU = "--%%%Result_Attendu:";

        private static string sqlFilePath;
        private static string excelFilePath;
        private static bool debug = false;
        private static int numberOfLinesToClear = 200;

        private static Dictionary<string, (string, string)> languageDictionary;
        private static Language selectedLanguage = Language.French; // Default to French

        public enum Language
        {
            French,
            English
        }

        
        static void Main()
        {
            InitializeLanguageDictionary(); // Initialize language dictionary with hardcoded values

            while (true)
            {
                try
                {
                    Console.WriteLine(GetLocalizedString("WelcomeMessage"));
                    Console.WriteLine(GetLocalizedString("ChooseOption"));
                    Console.WriteLine("1. " + GetLocalizedString("CustomFilePathOption"));
                    Console.WriteLine("2. " + GetLocalizedString("ExecuteOption"));
                    Console.WriteLine("3. " + GetLocalizedString("ToggleDebugOption"));
                    Console.WriteLine("4. " + GetLocalizedString("ModifyLinesToClearOption"));
                    Console.WriteLine("5. " + GetLocalizedString("ChooseLanguage"));
                    Console.WriteLine("0. " + GetLocalizedString("QuitOption"));

                    Console.Write(GetLocalizedString("ChoicePrompt"));
                    string choice = Console.ReadLine();

                    switch (choice)
                    {
                        case "1":
                            SetCustomFilePath();
                            break;
                        case "2":
                            PreExecution();
                            break;
                        case "3":
                            ToggleDebugMode();
                            break;
                        case "4":
                            ModifierNombreLigne();
                            break;
                        case "5":
                            ChooseLanguage();
                            break;
                        case "0":
                            AppQuit();
                            break;
                        default:
                            Console.WriteLine(GetLocalizedString("InvalidOption"));
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(GetLocalizedString("ErrorMessage") + ex.Message);
                }
            }
        }

        #region Ui logic
        private static void SetCustomFilePath()
        {
            // Ask for custom file paths
            Console.Write(GetLocalizedString("SqlFilePathPrompt"));
            sqlFilePath = Console.ReadLine();
            Console.Write(GetLocalizedString("ExcelFilePathPrompt"));
            excelFilePath = Console.ReadLine();
        }
        private static void PreExecution()
        {
            if (!string.IsNullOrEmpty(sqlFilePath) && !string.IsNullOrEmpty(excelFilePath))
            {
                if (sqlFilePath.EndsWith(".sql", StringComparison.OrdinalIgnoreCase) &&
                    excelFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    Execute();
                }
                else
                {
                    if (!sqlFilePath.EndsWith(".sql", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine(GetLocalizedString("SqlFileExtensionError"));
                    }
                    if (!excelFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine(GetLocalizedString("ExcelFileExtensionError"));
                    }
                }
            }
        }
        private static void ToggleDebugMode()
        {
            // Toggle debug mode
            debug = !debug;
            Console.WriteLine(debug ? GetLocalizedString("DebugModeActivated") : GetLocalizedString("DebugModeDeactivated"));
        }
        private static void ModifierNombreLigne()
        {
            // Modify number of lines to clear
            Console.Write(GetLocalizedString("ModifyLinesToClearOption"));
            if (int.TryParse(Console.ReadLine(), out int linesToClear) && linesToClear > 0)
            {
                numberOfLinesToClear = linesToClear;
                Console.WriteLine(GetLocalizedString("LinesToClearUpdated") + numberOfLinesToClear);
            }
            else
            {
                Console.WriteLine(GetLocalizedString("InvalidNumberOfLines"));
            }
        }
        private static void ChooseLanguage()
        {
            Console.WriteLine(GetLocalizedString("ChooseOption"));
            Console.WriteLine("1. Français");
            Console.WriteLine("2. English");

            Console.Write(GetLocalizedString("ChoicePrompt"));
            string langChoice = Console.ReadLine();

            switch (langChoice)
            {
                case "1":
                    selectedLanguage = Language.French;
                    break;
                case "2":
                    selectedLanguage = Language.English;
                    break;
                default:
                    Console.WriteLine(GetLocalizedString("InvalidOption"));
                    break;
            }
        }
        public static void AppQuit()
        {
            Console.WriteLine(GetLocalizedString("PressKeyToExit"));
            Console.ReadKey();
            Environment.Exit(0);
        }
        #endregion


        #region BackEnd
        private static void Execute()
        {
            try
            {
                Console.WriteLine(GetLocalizedString("ExecutionInProgress"));

                // Read the .sql file content
                string sqlContent = File.ReadAllText(sqlFilePath);

                // Split content based on '--%%%C_TabBalise'
                string[] sqlSegments = sqlContent.Split(new string[] { BALISE_TAB }, StringSplitOptions.RemoveEmptyEntries);

                // Trim each segment to remove leading and trailing whitespace/newlines
                List<string> trimmedSegments = sqlSegments.Select(segment => segment.Trim()).ToList();

                if (debug)
                {
                    // Print segments to console for debugging
                    foreach (var segment in trimmedSegments)
                    {
                        Console.WriteLine("Segment:");
                        Console.WriteLine(segment);
                        Console.WriteLine("----------");
                    }
                }

                // Now you can use trimmedSegments to write to Word or Excel
                WriteToExcel(trimmedSegments, excelFilePath);
                //WriteToWord(trimmedSegments);
            }
            catch (Exception ex)
            {
                Console.WriteLine(GetLocalizedString("ErrorMessage") + ex.Message);
                throw;
            }
        }
        static void WriteToExcel(List<string> segments, string excelPath)
        {
            // Set the license context (EPPlus version 5+)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                // Ensure the directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(excelPath));

                var fileInfo = new FileInfo(excelPath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    // Create a new worksheet if it does not already exist
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "SQL Data") ?? package.Workbook.Worksheets.Add("SQL Data");

                    ClearNNumberOfLine(numberOfLinesToClear, worksheet);

                    // Write each segment into the worksheet
                    for (int i = 0; i < segments.Count; i++)
                    {
                        string segment = segments[i];
                        string resultValue = "";
                        if (segment.Contains(RESULTAT_ATTENDU))
                        {
                            // Extract the value (1 or 0)
                            int startIndex = segment.IndexOf(RESULTAT_ATTENDU) + RESULTAT_ATTENDU.Length;
                            int endIndex = segment.IndexOfAny(new char[] { '\r', '\n' }, startIndex);
                            if (endIndex == -1)
                            {
                                endIndex = segment.Length;
                            }
                            resultValue = segment.Substring(startIndex, endIndex - startIndex).Trim();

                            // Remove the --%%%Result_Attendu: part from the segment
                            segment = segment.Remove(segment.IndexOf(RESULTAT_ATTENDU), endIndex - segment.IndexOf(RESULTAT_ATTENDU));
                        }

                        // Write the value to the next column
                        string sanitizedSegment = SanitizeSegment(segment);
                        worksheet.Cells[i + 3, 1].Value = sanitizedSegment;
                        worksheet.Cells[i + 3, 2].Value = resultValue;
                    }

                    package.Save();
                }
                Console.WriteLine(GetLocalizedString("ExcelFileCreated") + excelPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(GetLocalizedString("ErrorMessage") + ex.Message);
                throw;
            }
        }

        private static void ClearNNumberOfLine(int numberOfLinesToClear, OfficeOpenXml.ExcelWorksheet worksheet)
        {
            // Clear n number of lines before writing
            int start = 2;
            numberOfLinesToClear += start;
            for (int i = start; i < numberOfLinesToClear; i++)
            {
                worksheet.Cells[i + 1, 1, i + 1, worksheet.Dimension.End.Column].Clear();
            }
        }

        static void WriteToWord(List<string> segments)
        {
            Console.WriteLine("!! Error: WriteToWord is not implemented.");
            throw new NotImplementedException();
            //using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(@"path_to_your_output_word_file.docx", DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            //{
            //    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            //    mainPart.Document = new Document();
            //    Body body = new Body();

            //    foreach (var segment in segments)
            //    {
            //        Paragraph paragraph = new Paragraph(new Run(new Text(segment)));
            //        body.Append(paragraph);
            //    }

            //    mainPart.Document.Append(body);
            //    mainPart.Document.Save();
            //}
        }

        static string SanitizeSegment(string segment)
        {
            // Truncate the string if it exceeds the maximum length allowed by Excel
            int maxLengthExcel = 32767;
            if (segment.Length >= maxLengthExcel)
            {
                //segment = segment.Substring(0, maxLengthExcel - 1);
                string error = GetLocalizedString("SanitizeSegmentArgument");
                throw new ArgumentOutOfRangeException(string.Format(error, segment.Length, maxLengthExcel, segment));
            }
            return segment;
        }
        #endregion


        #region Langages
        private static string GetLocalizedString(string key)
        {
            // Get the localized string based on the current language.
            string localizedString;
            if (languageDictionary.TryGetValue(key, out var value))
            {
                switch (selectedLanguage)
                {
                    case Language.English:
                        localizedString = value.Item2;
                        break;
                    case Language.French:
                        localizedString = value.Item1;
                        break;
                    // Add more cases for other languages if needed
                    default:
                        localizedString = $"[Missing translation for '{key}']";
                        break;
                }
            }
            else
            {
                localizedString = $"[Missing translation for '{key}']";
            }
            return localizedString;
        }
        private static void InitializeLanguageDictionary()
        {
            // Initialize the language dictionary with hardcoded values for French and English
            languageDictionary = new Dictionary<string, (string, string)>
            {
                { "WelcomeMessage", ("Bienvenue dans l'application de conversion SQL vers Tableau!", "Welcome to SQL to Tableau conversion application!") },
                { "ChooseOption", ("Veuillez choisir une option:", "Please choose an option:") },
                { "CustomFilePathOption", ("Fournir des chemins de fichiers personnalisés", "Provide custom file paths") },
                { "ExecuteOption", ("Exécuter avec les fichiers", "Execute with files") },
                { "ToggleDebugOption", ("Activer/Désactiver le mode débogage", "Toggle debug mode") },
                { "ModifyLinesToClearOption", ("Modifier le nombre de lignes à effacer", "Modify lines to clear") },
                { "ChooseLanguage", ("Choisir langage", "Change language") },
                { "QuitOption", ("Quitter", "Quit") },
                { "ChoicePrompt", ("Choix: ", "Choice: ") },
                { "InvalidOption", ("Option invalide. Veuillez choisir une option valide.", "Invalid option. Please choose a valid option.") },
                { "ErrorMessage", ("Une erreur est survenue: ", "An error occurred: ") },
                { "SqlFilePathPrompt", ("Chemin du fichier SQL: ", "SQL file path: ") },
                { "ExcelFilePathPrompt", ("Chemin du fichier Excel: ", "Excel file path: ") },
                { "DebugModeActivated", ("Mode débogage activé", "Debug mode activated") },
                { "DebugModeDeactivated", ("Mode débogage désactivé", "Debug mode deactivated") },
                { "LinesToClearUpdated", ("Le nombre de lignes à effacer a été mis à jour: ", "Number of lines to clear updated: ") },
                { "InvalidNumberOfLines", ("Nombre de lignes invalide. Veuillez entrer un nombre supérieur à 0.", "Invalid number of lines. Please enter a number greater than 0.") },
                { "SqlFileExtensionError", ("Erreur: Le chemin du fichier SQL n'est pas valide.", "Error: SQL file path is not valid.") },
                { "ExcelFileExtensionError", ("Erreur: Le chemin du fichier Excel n'est pas valide.", "Error: Excel file path is not valid.") },
                { "ExecutionInProgress", ("Execution en cours...", "Execution in progress...") },
                { "PressKeyToExit", ("Appuyez sur une touche pour quitter...", "Press any key to exit...") },
                { "ExcelFileCreated", ("Fichier Excel créé à: ", "Excel file created at: ") },
                { "SanitizeSegmentArgument", (" --- EXCEPTION LANCÉE !! --- \n Le segment passé en paramètre est plus long que la longueur autorisée par Excel : {0} >= {1}. \n {2} \n Le segment passé en paramètre est plus long que la longueur autorisée par Excel : {0} >= {1}.", " --- EXCEPTION THROWN!! --- \n The segment passed as parameter is longer than the length allowed by Excel: {0} >= {1}. \n {2} \n The segment passed as parameter is longer than the length allowed by Excel: {0} >= {1}.") }
            };
        }
        #endregion
    }
}
