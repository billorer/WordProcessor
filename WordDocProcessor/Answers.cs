using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WordDocProcessor
{
    class Answers
    {
        private int seqParagraph = 1;
        private int seqListDecimal = 2;
        private int seqCircleList = 3;
        private int seqNotesList = 4;
        private int seqImages = 5;
        private int seqTables = 6;

        private List<string> listParagraph;

        private List<string> listDecimal; // nemartana ezeket matrixolni, ha esetlegesen tobb felsorolas van egy paragrafus reszben
        private List<string> listCircle;

        private List<string> listNotes;

        private List<string> listTables;

        private string finalAnswer;

        /// <summary>
        /// The Sequence of the elements
        /// 1. -> paragraph
        /// 2. -> decimalList
        /// 3. -> circleList
        /// 4. -> notesList
        /// 5. -> images
        /// 6. -> tables
        /// </summary>
        private List<int> sequence;


        public int SeqParagraph
        {
            get { return seqParagraph; }
        }
        public int SeqCircleList
        {
            get { return seqCircleList; }
        }
        public int SeqDecimalList
        {
            get { return seqListDecimal; }
        }
        public int SeqNotesList
        {
            get { return seqNotesList; }
        }
        public int SeqImage
        {
            get { return seqImages; }
        }
        public int SeqTables
        {
            get { return seqTables; }
        }

        public List<string> ListParagraph
        {
            get { return listParagraph; }
            set { listParagraph = value; }
        }

        public string SetListParagraphElement
        {
            set
            {
                listParagraph.Add(value);
            }
        }

        public List<string> ListDecimal
        {
            get { return listDecimal; }
            set { listDecimal = value; }
        }

        public string SetListDecimalElement
        {
            set {
                listDecimal.Add(value);    
            }
        }

        public List<string> ListCircle
        {
            get { return listCircle; }
            set { listCircle = value; }
        }

        public string SetListCircleElement
        {
            set
            {
                listCircle.Add(value);
            }
        }

        public List<string> ListNotes
        {
            get { return listNotes; }
            set { listNotes = value; }
        }

        public string SetListNotesElement
        {
            set
            {
                listNotes.Add(value);
            }
        }

        public List<string> ListTables
        {
            get { return listTables; }
            set { listTables = value; }
        }

        public string SetListTablesElement
        {
            set
            {
                listTables.Add(value);
            }
        }

        public List<int> Sequence
        {
            get { return sequence; }
            set { sequence = value; }
        }

        public int SetSequenceElement
        {
            set
            {
                sequence.Add(value);
            }
        }

        public string AnswerToString()
        {

            bool ulElement = false;

            for (int i = 0; i < sequence.Count; i++)
            {
                switch (sequence[i])
                {
                    case 1:
                        finalAnswer += "<p>" + listParagraph[0] + "</p>";
                        listParagraph.RemoveAt(0);
                        break;
                    case 3:
                        if (!ulElement)
                        {
                            finalAnswer += "<ul style='list-style-type:disc'>";
                            ulElement = true;
                        }

                        finalAnswer += "<li><p>";
                        finalAnswer += listCircle[0];
                        listCircle.RemoveAt(0);
                        finalAnswer += "</p></li>";

                        if (i + 1 < sequence.Count)
                        {
                            if (sequence[i + 1] != 3)
                            {
                                finalAnswer += "</ul>";
                                ulElement = false;
                            }
                        }

                        break;
                    case 2:
                        if (!ulElement)
                        {
                            finalAnswer += "<ul style='list-style-type:number'>";
                            ulElement = true;
                        }
                        

                        finalAnswer += "<li><p>";
                        finalAnswer += listDecimal[0];
                        listDecimal.RemoveAt(0);
                        finalAnswer += "</p></li>";

                        if (i + 1 < sequence.Count)
                        {
                            if (sequence[i + 1] != 2)
                            {
                                finalAnswer += "</ul>";
                            }
                        }

                        break;
                    case 4:
                        if (listNotes[0].Equals("Note"))
                        {
                            if(sequence[i - 1] == 4)
                            {
                                finalAnswer += "</div>";
                            }

                            finalAnswer += "<div class='boxWithNoSides'><p class='note'>";
                            finalAnswer += listNotes[0];
                            listNotes.RemoveAt(0);
                            finalAnswer += "</p>";
                        }
                        else
                        {
                            finalAnswer += listNotes[0];
                            listNotes.RemoveAt(0);
                        }

                        if (i + 1 < sequence.Count)
                        {
                            if (sequence[i + 1] != 4)
                            {
                                finalAnswer += "</div>";
                            }
                        }

                        break;
                    case 5:
                        string[] substrings = Regex.Split(listParagraph[0], @"/(.*)\\img_.*(.*)/");
                        int imageNumber = Int32.Parse(Regex.Match(substrings[0], @"\d+").Value);                       
                        string[] modifiedBeforeText = Regex.Split(substrings[0], "img_" + imageNumber);

                        if (modifiedBeforeText[0] != "")
                        {
                            finalAnswer += "<p>" + modifiedBeforeText[0];
                            finalAnswer += "<img src= '/Home/GetImage?imgNumber=" + imageNumber + "'/>";
                        }
                        else
                        {
                            finalAnswer += "<img class='figureImage' src= '/Home/GetImage?imgNumber=" + imageNumber + "'/>";
                        }
                        
                        if(modifiedBeforeText[1] != "")
                        {
                            finalAnswer += modifiedBeforeText[1]
                                + "</p>";
                        }

                        listParagraph.RemoveAt(0);
                        break;
                    case 6:
                        finalAnswer += listTables[0];
                        listTables.RemoveAt(0);
                        break;
                }
            }
            return finalAnswer;
        }

        public Answers()
        {
            listParagraph = new List<string>();    
            listDecimal = new List<string>();
            listCircle = new List<string>();
            listNotes = new List<string>();
            listTables = new List<string>();

            sequence = new List<int>();

            finalAnswer = "";
        }
    }
}
