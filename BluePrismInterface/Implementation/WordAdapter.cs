using OfficeExtension;
using BluePrismInterface.Interfaces;
using System.Data;
using System;

namespace BluePrismInterface.Implementations
{
    public class WordAdapter : IBluePrismAdapter
    {
        private string _documentFilePath;
        private string _IMAGE_FILE_PATH_LABEL = "FileImage";
        private string _TABLE_INDEX_LABEL = "TableIndex";
        private string _COLUMN_INDEX_LABEL = "Column";
        private string _ROW_INDEX_LABEL = "Row";
        private string _SUBTITLE_LABEL = "Texto";
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

                    if (Int32.TryParse(row[_TABLE_INDEX_LABEL].ToString(), out int tableIndex))
                    {
                        DocClass.AppendImageOnTableColumn(imageFilePath, tableIndex, _INSERT_COLUMN_INDEX);
                        DocClass.AppendTextOnTableColumn(imageSubtitle, tableIndex, _INSERT_COLUMN_INDEX);
                    }
                    else
                        throw new Exception("err");
                }

                DocClass.SaveDocAs(this._documentFilePath);
            }
        }

        public void InserTextIntoWordTableFromDataTable(DataTable datatable)
        {
            using (WordDocument DocClass = new WordDocument())
            {
                DocClass.OpenFile(this._documentFilePath);

                foreach (DataRow row in datatable.Rows)
                {
                    string text = row.Field<string>(_TEXT_LABEL);
                    int tableIndex = Convert.ToInt32(row.Field<int>(_TABLE_INDEX_LABEL));
                    int columnIndex = Convert.ToInt32(row.Field<int>(_COLUMN_INDEX_LABEL));
                    int rowIndex = Convert.ToInt32(row.Field<int>(_ROW_INDEX_LABEL));

                    DocClass.InsertTextOnTableCell(text, tableIndex, rowIndex, columnIndex);
                }

                DocClass.SaveDocAs(this._documentFilePath);
            }
        }
    }
}
