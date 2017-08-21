using Microsoft.Office.Interop.Word;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;

using System.Windows.Forms;
using System.Drawing;

namespace WordDocProcessor
{
    class Program
    {
        private static Microsoft.Office.Interop.Word.Application word;
        private static Document docs;

        /// <summary>
        /// The list in which we keep the questions
        /// </summary>
        private static List<string> questions;

        /// <summary>
        /// The list in which we keep the answers
        /// </summary>
        private static List<Answers> answersList;

        private static Answers currentAnswer;

        private static List<string> docTablesList;

        private static int imageNumber = 1;
        private static int tableNumber = 0;

        /// <summary>
        /// This function creates a WORD DOCUMENT object and returns it.
        /// It Opens the adequate word document.
        /// </summary>
        /// <returns></returns>
        private static Document GetDocument(){
            word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object path = @"c:\Users\zoran.suto\Documents\OMG3.docx";
            object readOnly = true;
            docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            return docs;  
        }
        private static string stripCellText(string text)
        {
            return text.Replace("\r", "");
        }
        static void Main(string[] args)
        {
            // Initializing the list variables
            questions = new List<string>();
            answersList = new List<Answers>();
            currentAnswer = new Answers();
            docTablesList = new List<string>();

            Document docs = GetDocument();

            string heading2 = "";
            string heading3 = "";
            string heading4 = "";
            string previousParentHeading = "";

            bool foundFirstHeading = false;
            bool signTable = false;          

            ExtractImagesFromDocIntoWordDocImages();

            

            foreach (Table tb in docs.Tables)
            {
                string table = "<table>";
                for (int row = 1; row <= tb.Rows.Count; row++)
                {
                    table += "<tr>";
                    for (int column = 1; column <= tb.Columns.Count; column++)
                    {
                        if(column == 1)
                        {
                            table += "<th>";
                        }
                        else
                        {
                            table += "<td>";
                        }

                        try
                        {
                            var text = stripCellText(tb.Cell(row, column).Range.Text);
                            text = RemoveNewLineFromString(text);
                            text = CleanInvalidXmlChars(text);
                            table += text;
                            //Console.WriteLine("Row: " + row + " Column: " + column + " Text: " + text);
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {

                        }

                        if (column == 1)
                        {
                            table += "</th>";
                        }
                        else
                        {
                            table += "</td>";
                        }
                    }
                    table += "</tr>";
                }
                table += "</table>";
                docTablesList.Add(table);
            }



            ExportDataToXMLFile("tables","table", @"c:\tables.xml", docTablesList);


            //docTablesList = XDocument.Load(@"c:\tables.xml").Root.Elements("table")
            //               .Select(element => element.Value)
            //               .ToList();

            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                Style style = docs.Paragraphs[i + 1].get_Style() as Style;
                string styleName = style.NameLocal;
                string currentText = docs.Paragraphs[i + 1].Range.Text.ToString();

                // We remove the inproper characters
                currentText = RemoveNewLineFromString(currentText);
                currentText = CleanInvalidXmlChars(currentText);

                // Just to get minimal information about the program when it runs
                if (i % 100 == 0)
                {
                    Console.WriteLine(currentText);
                }
                //Console.WriteLine("Style: " + styleName + " Text: " + currentText);

                // From now on we can start adding questions and their answers
                if(styleName == "Heading 2")
                {
                    foundFirstHeading = true;
                }

                if (foundFirstHeading)
                {
                    if (styleName == "Heading 2")
                    {
                        heading2 = currentText;
                        previousParentHeading = heading2;
                        questions.Add(heading2);

                        AddAnswerAndCreateNewOne();

                    }
                    else if (styleName == "Heading 3")
                    {
                        heading3 = heading2 + " " + currentText;
                        previousParentHeading = heading3;

                        questions.Add(heading3);
                        AddAnswerAndCreateNewOne();

                    }
                    else if (styleName == "Heading 4")
                    {
                        heading4 = heading3 + " " + currentText;
                        previousParentHeading = heading4;

                        questions.Add(heading4);
                        AddAnswerAndCreateNewOne();

                    }
                    else if (styleName == "Titel:Label")
                    {
                        questions.Add(previousParentHeading + " " + currentText);
                        AddAnswerAndCreateNewOne();
                    }
                    else if (styleName == "Liste:Punkt")
                    {
                        currentAnswer.SetListCircleElement = currentText;
                        currentAnswer.SetSequenceElement = currentAnswer.SeqCircleList;
                    }
                    else if (styleName == "Liste:Nummer")
                    {
                        currentAnswer.SetListDecimalElement = currentText;
                        currentAnswer.SetSequenceElement = currentAnswer.SeqDecimalList;
                    }
                    else if (styleName == ":imp")
                    {
                        currentAnswer.SetListNotesElement = currentText;
                        currentAnswer.SetSequenceElement = currentAnswer.SeqNotesList;
                    }
                    else if (styleName.Contains("Tab:Kopf"))
                    {
                        //currentAnswer.SetListNotesElement = currentText;
                        if (!signTable)
                        {
                            currentAnswer.SetSequenceElement = currentAnswer.SeqTables;
                            currentAnswer.SetListTablesElement = docTablesList[tableNumber++];
                        }
                        
                        signTable = true;
                    }
                    else if (styleName.Contains("Tab:Absatz"))
                    {
                        //currentAnswer.SetListNotesElement = currentText;
                        //currentAnswer.SetSequenceElement = currentAnswer.SeqNotesList;
                        signTable = false;
                    }
                    else if (currentText == "/")
                    {
                        currentAnswer.SetListParagraphElement = currentText.Replace("/", "img_" + imageNumber.ToString());
                        imageNumber++;
                        currentAnswer.SetSequenceElement = currentAnswer.SeqImage;
                    }
                    else if (currentText != "")
                    {
                        currentAnswer.SetListParagraphElement = currentText;
                        currentAnswer.SetSequenceElement = currentAnswer.SeqParagraph;
                    }

                    if (i == docs.Paragraphs.Count - 1) //the last questions answer
                    {
                        AddAnswerAndCreateNewOne();
                    }
                }
            }

            docs.Close();
            word.Quit();

            ExportDataToXMLFile("answers", "answer", @"c:\answers.xml", answersList);
            ExportDataToXMLFile("questions", "question", @"c:\questions.xml", questions);

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        private static void AddAnswerAndCreateNewOne()
        {
            if(currentAnswer.Sequence.Count != 0)
            {
                answersList.Add(currentAnswer);
                currentAnswer = new Answers();
            }
        }

        private static void ExtractImagesFromDocIntoWordDocImages()
        {
            for (var index = 1; index <= word.ActiveDocument.InlineShapes.Count; index++)
            {
                var inlineShapeId = index;

                // parameterized thread start
                var thread = new Thread(() => SaveInlineShapeToFile(inlineShapeId));

                // STA is needed in order to access the clipboard
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            }
        }

        protected static void SaveInlineShapeToFile(int inlineShapeId)
        {
            // Get the shape, select, and copy it to the clipboard
            var inlineShape = docs.InlineShapes[inlineShapeId];
            inlineShape.Select();
            docs.ActiveWindow.Selection.CopyAsPicture();

            // Check data is in the clipboard
            if (Clipboard.GetDataObject() != null)
            {
                var data = Clipboard.GetDataObject();

                // Check if the data conforms to a bitmap format
                if (data != null && data.GetDataPresent(DataFormats.Bitmap))
                {
                    // Fetch the image and convert it to a Bitmap
                    Image image = (Image)data.GetData(DataFormats.Bitmap, true);

                    if (image != null)
                    {
                        var currentBitmap = new Bitmap(image);
                        System.IO.Directory.CreateDirectory(@"c:\WordDocImages\");
                        // Save the bitmap to a file
                        currentBitmap.Save(@"c:\WordDocImages\" + String.Format("img_{0}.png", inlineShapeId));
                    }
                }
}
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

        private static void ExportDataToXMLFile(string rootElementName, string childElementName, string filePath, List<Answers> data)
        {
            new XElement("answers",
                data.Select(answerElement => new XElement("answer",
                    new XElement("paragraphList", answerElement.ListParagraph.Select(paragraphList =>
                        new XElement("paragraphListElement", paragraphList
                    ))),
                    new XElement("circleList", answerElement.ListCircle.Select(circleList => 
                        new XElement("circleListElement", circleList
                    ))),
                    new XElement("decimalList", answerElement.ListDecimal.Select(decimalList =>
                        new XElement("decimalListElement", decimalList
                    ))),
                    new XElement("notesList", answerElement.ListNotes.Select(noteList =>
                        new XElement("noteListElement", noteList
                    ))),
                    new XElement("tableList", answerElement.ListTables.Select(tableList =>
                        new XElement("table", tableList
                    ))),
                    new XElement("sequence", answerElement.Sequence.Select(sequence =>
                        new XElement("sequenceElement", sequence
                    )))
                ))).Save(filePath, SaveOptions.OmitDuplicateNamespaces);
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
