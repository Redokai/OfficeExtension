using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace OfficeExtension
{
    public class WordDocument : IDisposable
    {
        private string _NEW_ROW_TEXT_CONTENT = "\r\a";

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
            this._doc = this._app.Documents.Open(filePath, Visible: false, ReadOnly: false);
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
            Range insertRange = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            if (insertRange == null)
            {
                _AppendTableRow(table);
            }
            insertRange = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            _InsertPictureInRange(imagePath, insertRange);
        }

        public void AppendImageOnTableColumn(string imagePath, int tableIndex, int columnIndex)
        {
            Table table = _FindTable(tableIndex);
            Range insertRange = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
            if (insertRange == null)
            {
                _AppendTableRow(table);
            }
            insertRange = _FindTableEmptyCellOnSpecificCollumn(table, columnIndex);
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

        public void AppendTextOnTableColumn(string text, int tableIndex, int columnIndex)
        {
            Table table = _FindTable(tableIndex);
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

        private Table _FindTable(int tableIndex)
        {
            return this._doc.Tables[tableIndex];
        }

        private Range _FindTableEmptyCellOnSpecificCollumn(Table table, int columnIndex)
        {
            for (int i = 1; i < table.Rows.Count + 1; i++)
            {
                Cell cell = table.Cell(i, columnIndex);
                if (cell.Range.Text == _NEW_ROW_TEXT_CONTENT && cell.Range.InlineShapes.Count == 0)
                {
                    return cell.Range;
                }
            }
            return null;
        }

        private void _AppendTableRow(Table table)
        {
            Object oMissing = System.Reflection.Missing.Value;
            table.Rows.Add(ref oMissing);
        }

        private void _InsertPictureInRange(string path, Range range)
        {
            range.InlineShapes.AddPicture(path);
        }
        private void _InsertTextInRange(string text, Range range)
        {
            range.Text = text;
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                }
                this.Close();
                this.Quit();

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion


    }


}
