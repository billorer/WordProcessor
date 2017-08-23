using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace WordDocProcessor
{
    class Program
    {
        private static Application word;
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
        private static List<string> formattedAnswersList;

        private static List<string> docTablesList;

        private static int imageNumber = 1;
        private static int tableNumber = 1;

        /// <summary>
        /// This function creates a WORD DOCUMENT object and returns it.
        /// It Opens the adequate word document.
        /// </summary>
        /// <returns></returns>
        private static Document GetDocument(){
            word = new Application();
            object miss = System.Reflection.Missing.Value;
            object path = PublicFunctionsVariables.wordDocumentFilePath;
            object readOnly = true;
            docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            return docs;  
        }

        static void Main(string[] args)
        {
            // Initializing the list variables
            questions = new List<string>();
            answersList = new List<Answers>();
            currentAnswer = new Answers();
            docTablesList = new List<string>();
            formattedAnswersList = new List<string>();

            Document docs = GetDocument();

            string heading2 = "";
            string heading3 = "";
            string heading4 = "";
            string previousParentHeading = "";

            bool foundFirstHeading = false;
            bool signTable = false;

            ExtractObjectsFromWord.ExtractImagesFromDocIntoFile(word, docs, PublicFunctionsVariables.imagesFolderPath);

            docTablesList = ExtractObjectsFromWord.ExtractTablesFromDocIntoList(docs);
            PublicFunctionsVariables.ExportDataToXMLFile("tables","table", @"c:\tables.xml", docTablesList);
            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                Style style = docs.Paragraphs[i + 1].get_Style() as Style;
                string styleName = style.NameLocal;
                string currentText = docs.Paragraphs[i + 1].Range.Text.ToString();

                // We remove the inproper characters
                currentText = PublicFunctionsVariables.RemoveNewLineFromString(currentText);
                currentText = PublicFunctionsVariables.CleanInvalidXmlChars(currentText);

                // Just to get minimal information about the program when it runs
                if (i % 100 == 0)
                {
                    Console.WriteLine(currentText);
                }
                //Console.WriteLine("Style: " + styleName + " Text: " + currentText);

                // From now on we can start adding questions and their answers
                if (styleName == "Heading 2")
                {
                    foundFirstHeading = true;
                }

                // This is how currently we try to figure it out when a table ends...
                if (!styleName.Contains("Tab:Kopf") && !styleName.Contains("Tab:Absatz") && !styleName.Contains("Normal"))
                {
                    signTable = false;
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

                        if (AddAnswerAndCreateNewOne())
                        {
                            questions.Add(heading3);
                        }
                    }
                    else if (styleName == "Heading 4")
                    {
                        heading4 = heading3 + " " + currentText;
                        previousParentHeading = heading4;


                        if (AddAnswerAndCreateNewOne())
                        {
                            questions.Add(heading4);
                        }

                    }
                    else if (styleName == "Titel:Label")
                    {
                        if (AddAnswerAndCreateNewOne())
                        {
                            questions.Add(previousParentHeading + " " + currentText);
                        }
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
                        if (!signTable)
                        {
                            currentAnswer.SetSequenceElement = currentAnswer.SeqTables;
                            currentAnswer.SetListTablesElement = docTablesList[tableNumber];
                            tableNumber++;
                            signTable = true;
                        }

                        
                    }
                    else if (styleName.Contains("Tab:Absatz"))
                    {
                        //signTable = false;
                    }
                    else if (currentText == "/")
                    {
                        currentAnswer.SetListParagraphElement = currentText.Replace("/", "img_" + imageNumber.ToString());
                        imageNumber++;
                        currentAnswer.SetSequenceElement = currentAnswer.SeqImage;
                    }
                    else if (currentText.Contains(" /"))
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

            PublicFunctionsVariables.ExportDataToXMLFile("questions", "question", PublicFunctionsVariables.questionsFilePath, questions);

            foreach (Answers ans in answersList)
            {
                formattedAnswersList.Add(ans.AnswerToString());
            }

            PublicFunctionsVariables.ExportDataToXMLFile("answers", "answer", PublicFunctionsVariables.answersFilePath, formattedAnswersList);

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

       
        private static bool AddAnswerAndCreateNewOne()
        {
            if(currentAnswer.Sequence.Count != 0)
            {
                answersList.Add(currentAnswer);
                currentAnswer = new Answers();
                return true;
            }
            return false;
        }

        /// <summary>
        /// This function exports the ANSWERS list as an XML file
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="data"></param>
        private static void ExportAnswersDataToXMLFile(string filePath, List<Answers> data)
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
    }
}
