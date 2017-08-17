using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordDocProcessor
{
    class Program
    {
        private static Microsoft.Office.Interop.Word.Application word;
        private static Microsoft.Office.Interop.Word.Document docs;

        /// <summary>
        /// The list in which we keep the questions
        /// </summary>
        private static List<string> questions;

        /// <summary>
        /// The list in which we keep the answers
        /// </summary>
        private static List<string> answers;

        /// <summary>
        /// This function creates a WORD DOCUMENT object and returns it.
        /// It Opens the adequate word document.
        /// </summary>
        /// <returns></returns>
        private static Microsoft.Office.Interop.Word.Document GetDocument(){
            word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object path = @"c:\Users\zoran.suto\Documents\OMG.docx";
            object readOnly = true;
            docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            return docs;  
        }

        static void Main(string[] args)
        {
            // Initializing the list variables
            questions = new List<string>();
            answers = new List<string>();

            Microsoft.Office.Interop.Word.Document docs = GetDocument();

            string answer = "";
            string heading2 = "";
            string heading3 = "";
            string heading4 = "";
            string previousParentHeading = "";

            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {


                Microsoft.Office.Interop.Word.Style style = docs.Paragraphs[i + 1].get_Style() as Microsoft.Office.Interop.Word.Style;
                string styleName = style.NameLocal;
                string currentText = docs.Paragraphs[i + 1].Range.Text.ToString();

                // We remove the inproper characters
                currentText = RemoveNewLineFromString(currentText);
                currentText = CleanInvalidXmlChars(currentText);

                //if (i % 100 == 0)
                //{
                //    Console.WriteLine(currentText);
                //}

                if (styleName == "Heading 2")
                {
                    heading2 = currentText;
                    previousParentHeading = heading2;
                    questions.Add(heading2);
                    if (questions.Count != 0 && answer != "")
                    {                       
                        answers.Add(answer);
                        answer = "";
                    }
                }
                else if (styleName == "Heading 3")
                {
                    heading3 = heading2 + " " + currentText;
                    previousParentHeading = heading3;
                    
                    if (questions.Count != 0 && answer != "")
                    {
                        questions.Add(heading3);
                        answers.Add(answer);
                        answer = "";
                    }
                }
                else if (styleName == "Heading 4")
                {
                    heading4 = heading3 + " " + currentText;
                    previousParentHeading = heading4;
                    
                    if (questions.Count != 0 && answer != "")
                    {
                        questions.Add(heading4);
                        answers.Add(answer);
                        answer = "";
                    }
                }else if(styleName == "Titel:Label")
                {
                    
                    if (questions.Count != 0 && answer != "")
                    {
                        questions.Add(previousParentHeading + " " + currentText);
                        answers.Add(answer);
                        answer = "";
                    }
                }
                else if (!currentText.Contains("Figure") || currentText != " / ") //the currentText should contain only the answer section
                {
                    if (questions.Count != 0)
                    {
                        answer += currentText + " ";
                    }
                }
                else
                {
                    Console.WriteLine("Style: " + styleName + " Text: " + currentText);
                }
                if (i == docs.Paragraphs.Count - 1) //the last questions answer
                {
                    answers.Add(answer);
                    answer = "";
                }
            }

            docs.Close();
            word.Quit();

            ExportDataToXMLFile("answers", "answer", @"c:\answers.xml", answers);
            ExportDataToXMLFile("questions", "question", @"c:\questions.xml", questions);

            Console.WriteLine("Press any key to exit.");
            System.Console.ReadKey();
        }

        /// <summary>
        /// This function removes invalid XML characters from a string and returns the result
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string CleanInvalidXmlChars(string text)
        {
            // From xml spec valid chars: 
            // #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]     
            // any Unicode character, excluding the surrogate blocks, FFFE, and FFFF. 
            string re = @"[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000-x10FFFF]";
            return Regex.Replace(text, re, "");
        }

        /// <summary>
        /// This function exports a list of strings to a file in XML format
        /// </summary>
        /// <param name="rootElementName"></param>
        /// <param name="childElementName"></param>
        /// <param name="filePath"></param>
        private static void ExportDataToXMLFile(string rootElementName, string childElementName, string filePath, List<string> data)
        {
            new XElement(rootElementName, data.Select(i => new XElement(childElementName, i)))
                .Save(filePath, SaveOptions.OmitDuplicateNamespaces);
        }

        /// <summary>
        /// This function exports a list of strings to file
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        private static void ExportDataToFile(string fileName, List<string> data)
        {
            using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(fileName))
            {
                foreach (string line in data)
                {
                    file.WriteLine(line);
                }
            }
        }

        /// <summary>
        /// This function removes the new line from the prameter string and returns the new string
        /// </summary>
        /// <param name="line">The line from which the new line should be removed</param>
        /// <returns></returns>
        private static string RemoveNewLineFromString(string line)
        {
            line = line.TrimStart('\a');
            return line.TrimEnd('\r', '\n', '\a');
        }
    }
}
