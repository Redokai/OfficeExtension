using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace OfficeExtension
{
    public class WordDocument
    {
        private string _NEW_ROW_TEXT_CONTENT = "\r\a";
        private float _PAGE_WIDTH = 424.68F;

        private Application _app;
        private Document _doc;

        public WordDocument()
        {
            this._doc = null;
            this._app = null;
        }

        public Document OpenFile(string filePath)
        {
            this._app = new Application();
            this._doc = this._app.Documents.Open(filePath, Visible: true, ReadOnly: true);
            this._doc.Activate();
            this._doc.ActiveWindow.View.ReadingLayout = false;
            return _doc;
        }

        public void SaveDocAs(string filepath)
        {
            this._doc.SaveAs2(filepath);
        }
        public void Close()
        {
            this._doc.Close(SaveChanges: false);
        }
        public void Quit()
        {
            this._app.Quit(SaveChanges: false);
        }

        public void AppendImageOnTableColumn(string imagePath, string tableTitle, int columnIndex)
        {
            Table table = _FindTable(tableTitle);
            Cell insertCell = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            if (insertCell == null)
            {
                _AppendTableRow(table);
            }
            insertCell = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            _InsertPictureInCell(imagePath, insertCell);
        }

        public void AppendTextOnTableColumn(string text, string tableTitle, int columnIndex)
        {
            Table table = _FindTable(tableTitle);
            Cell insertCell = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            if (insertCell == null)
            {
                _AppendTableRow(table);
            }
            insertCell = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            _InsertTextInCell(text, insertCell);
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

        private Cell _FindTableEmptyCellOnSpecificCollumn(Table table, int columnIndex)
        {
            for (int i = 1; i < table.Rows.Count + 1; i++)
            {
                Cell cell = table.Cell(i, columnIndex);
                if (cell.Range.Text == _NEW_ROW_TEXT_CONTENT && cell.Range.InlineShapes.Count == 0)
                {
                    return cell;
                }
            }
            return null;
        }

        private void _AppendTableRow(Table table)
        {
            Object oMissing = System.Reflection.Missing.Value;
            table.Rows.Add(ref oMissing);
        }

        private void _InsertPictureInCell(string path, Cell cell)
        {
            InlineShape shape = cell.Range.InlineShapes.AddPicture(path);
            _ResizeImageToPageWidth(shape);
        }

        private void _ResizeImageToPageWidth(InlineShape shape)
        {
            shape.Width = 424.68F;
        }

        private void _InsertTextInCell(string text, Cell cell)
        {
            cell.Range.Text = text;
            _ScaleCellHeight(3, cell);
        }

        private void _ScaleCellHeight(float scale,Cell cell)
        {
            cell.Height *= scale; 
        }


    }


}
