using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordDocProcessor
{
    public class PublicFunctionsVariables
    {
        public static string questionsFilePath = @"c:\questions.xml";
        public static string answersFilePath = @"c:\answers.xml";
        public static string imagesFolderPath = @"c:\WordDocImages\";
        public static string wordDocumentFilePath = @"c:\Users\zoran.suto\Documents\User Manual.doc";

        /// <summary>
        /// this function removes carriage return from a string and returns it
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string StripCellText(string text)
        {
            return text.Replace("\r", "");
        }

        /// <summary>
        /// This function removes the new line from the prameter string and returns the new string
        /// </summary>
        /// <param name="line">The line from which the new line should be removed</param>
        /// <returns></returns>
        public static string RemoveNewLineFromString(string line)
        {
            line = line.TrimStart('\a');
            return line.TrimEnd('\r', '\n', '\a');
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
        /// This function exports a list of strings to file
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        public static void ExportDataToFile(string fileName, List<string> data)
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
        /// This function exports a list of strings to a file in XML format
        /// </summary>
        /// <param name="rootElementName"></param>
        /// <param name="childElementName"></param>
        /// <param name="filePath"></param>
        public static void ExportDataToXMLFile(string rootElementName, string childElementName, string filePath, List<string> data)
        {
            new XElement(rootElementName, data.Select(i => new XElement(childElementName, i)))
                .Save(filePath, SaveOptions.OmitDuplicateNamespaces);
        }
    }
}
