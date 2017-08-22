using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace WordDocProcessor
{
    public class ExtractObjectsFromWord
    {
        private static string fileLocation;

        /// <summary>
        /// This function extracts the images from the given document and saves them into a folder given
        /// </summary>
        /// <param name="word"></param>
        /// <param name="docs"></param>
        /// <param name="filePath"></param>
        public static void ExtractImagesFromDocIntoFile(Microsoft.Office.Interop.Word.Application word, Document docs, string filePath)
        {
            fileLocation = filePath;

            for (var index = 1; index <= word.ActiveDocument.InlineShapes.Count; index++)
            {
                var inlineShapeId = index;

                // parameterized thread start
                var thread = new Thread(() => SaveInlineShapeToFile(inlineShapeId, docs));

                // STA is needed in order to access the clipboard
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            }
        }

        private static void SaveInlineShapeToFile(int inlineShapeId, Document docs)
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
                        System.IO.Directory.CreateDirectory(fileLocation);
                        // Save the bitmap to a file
                        currentBitmap.Save(fileLocation + String.Format("img_{0}.png", inlineShapeId));
                    }
                }
            }
        }

        /// <summary>
        /// This function creates a list, which has all the tables from the document and returns it
        /// It also makes the html code around the table
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public static List<string> ExtractTablesFromDocIntoList(Document docs)
        {
            List<string> docTablesList = new List<string>();

            foreach (Table tb in docs.Tables)
            {
                string table = "<table>";
                for (int row = 1; row <= tb.Rows.Count; row++)
                {
                    table += "<tr>";
                    for (int column = 1; column <= tb.Columns.Count; column++)
                    {
                        string cellText = "";
                        try
                        {
                            cellText = PublicFunctionsVariables.StripCellText(tb.Cell(row, column).Range.Text);
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {

                        }

                        if (row == 1)
                        {
                            table += "<th>";
                        }
                        else
                        {
                            table += "<td>";
                        }

                        cellText = PublicFunctionsVariables.RemoveNewLineFromString(cellText);
                        cellText = PublicFunctionsVariables.CleanInvalidXmlChars(cellText);
                        table += cellText;

                        if (row == 1)
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
            return docTablesList;
        }
    }
}
