using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public Answers()
        {
            listParagraph = new List<string>();    
            listDecimal = new List<string>();
            listCircle = new List<string>();
            listNotes = new List<string>();
            listTables = new List<string>();

            sequence = new List<int>();
        }
    }
}
