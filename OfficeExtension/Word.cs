using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace OfficeExtension
{
    public class WordDocument
    {
        private Document _doc;

        public WordDocument()
        {
            this._doc = null;
        }

        public Document CreateFile()
        {
            Application app = new Application();
            this._doc = app.Documents.Add();
            return _doc;
        }

        public Document OpenFile(string filePath)
        {
            Application app = new Application();
            this._doc = app.Documents.Open(filePath, Visible: false, ReadOnly: true);
            return _doc;
        }

        public void SaveDocAs(string filepath)
        {
            this._doc.SaveAs2(filepath);
        }


        public void AppendImageOnTableColumn(string imagePath, string tableTitle, int columnIndex)
        {
            Table table = _FindTable(tableTitle);
            Range insertRange = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            if (insertRange == null)
            {

            }
            _InsertPictureInRange(imagePath, insertRange);
        }

        public void AppendTextOnTableColumn(string text, string tableTitle, int columnIndex)
        {
            Table table = _FindTable(tableTitle);
            Range insertRange = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            if (insertRange == null)
            {
                _AppendTableRow(table);
            }
            insertRange = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            _InsertTextInRange(text, insertRange);
        }

        private Table _FindTable(string tableTitle)
        {
            foreach (Table table in this._doc.Tables)
            {
                if (table.Title == tableTitle)
                    return table;
            }
            return null;
        }

        private Range _FindTableEmptyCellOnSpecificCollumn(Table table, int columnIndex)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                Cell cell = table.Cell(i, columnIndex);
                if (cell.Range.Text != null && cell.Range.InlineShapes.Count == 0)
                {
                    return cell.Range;
                }
            }
            return null;
        }

        private void _AppendTableRow(Table table)
        {
            table.Rows.Add();
        }

        private void _InsertPictureInRange(string path, Range range)
        {
            range.InlineShapes.AddPicture(path);
        }
        private void _InsertTextInRange(string text, Range range)
        {
            range.Text = text;
        }
    }


}
