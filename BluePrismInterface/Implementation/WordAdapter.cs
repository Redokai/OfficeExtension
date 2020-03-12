using OfficeExtension;
using BluePrismInterface.Interfaces;
using System.Data;
using System;

namespace BluePrismInterface.Implementations
{
    public class WordAdapter : IBluePrismAdapter, IDisposable
    {
        private string _documentFilePath;
        private string _IMAGE_FILE_PATH_LABEL = "FileImage";
        private string _TABLE_INDEX_LABEL = "TableIndex";
        private string _COLUMN_INDEX_LABEL = "Column";
        private string _ROW_INDEX_LABEL = "Row";
        private string _SUBTITLE_LABEL = "Texto";
        private string _TOKEN_LABEL = "Token";
        private string _TEXT_LABEL = "Text";
        private int _INSERT_COLUMN_INDEX = 1;

        public WordAdapter(string documentFilePath)
        {
            this._documentFilePath = documentFilePath;
        }

        public void InsertImagesIntoWordFromDataTable(DataTable datatable)
        {
            using (WordDocument DocClass = new WordDocument())
            {
                DocClass.OpenFile(this._documentFilePath);

                foreach (DataRow row in datatable.Rows)
                {
                    string imageFilePath = row.Field<string>(_IMAGE_FILE_PATH_LABEL);
                    string imageSubtitle = row.Field<string>(_SUBTITLE_LABEL);
                    int tableIndex = Decimal.ToInt32((Decimal)row[_TABLE_INDEX_LABEL]);

                    DocClass.AppendImageOnTableColumn(imageFilePath, tableIndex, _INSERT_COLUMN_INDEX);
                    DocClass.AppendTextOnTableColumn(imageSubtitle, tableIndex, _INSERT_COLUMN_INDEX);
                }

                DocClass.SaveDocAs(this._documentFilePath);
            }
        }

        public void InserTextIntoWordTableFromDataTable(DataTable datatable)
        {
            try
            {
                using (WordDocument DocClass = new WordDocument())
                {
                    DocClass.OpenFile(this._documentFilePath);

                    foreach (DataRow row in datatable.Rows)
                    {
                        string text = row.Field<string>(_TEXT_LABEL);
                        int tableIndex = Decimal.ToInt32((Decimal)row[_TABLE_INDEX_LABEL]);
                        int columnIndex = Decimal.ToInt32((Decimal)row[_COLUMN_INDEX_LABEL]);
                        int rowIndex = Decimal.ToInt32((Decimal)row[_ROW_INDEX_LABEL]);

                        DocClass.InsertTextOnTableCell(text, tableIndex, rowIndex, columnIndex);
                    }

                    DocClass.SaveDocAs(this._documentFilePath);
                }
            }
            catch (Exception err)
            {
                throw new Exception(err.ToString());
            }

        }

        public void ReplaceTexts(DataTable datatable)
        {
            try
            {
                using (WordDocument DocClass = new WordDocument())
                {
                    DocClass.OpenFile(this._documentFilePath);

                    foreach (DataRow row in datatable.Rows)
                    {
                        string text = row.Field<string>(_TEXT_LABEL);
                        string token = row.Field<string>(_TOKEN_LABEL);

                        DocClass.ReplaceTokenByText(token, text);
                    }

                    DocClass.SaveDocAs(this._documentFilePath);
                }
            }
            catch (Exception err)
            {
                throw new Exception(err.ToString());
            }
        }

        public void FindAndReplace(string token, string text)
        {
            try
            {
                using (WordDocument DocClass = new WordDocument())
                {
                    DocClass.OpenFile(this._documentFilePath);
                    DocClass.ReplaceTokenByText(token, text);
                    DocClass.SaveDocAs(this._documentFilePath);
                }
            }
            catch (Exception err)
            {
                throw new Exception(err.ToString());
            }
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }
}
