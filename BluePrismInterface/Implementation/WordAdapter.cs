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
        private string _SUBTITLE_LABEL = "Texto";
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
                    int tableIndex = row.Field<int>(_TABLE_INDEX_LABEL);
                    string imageSubtitle = row.Field<string>(_SUBTITLE_LABEL);

                    DocClass.AppendImageOnTableColumn(imageFilePath, tableIndex, _INSERT_COLUMN_INDEX);
                    DocClass.AppendTextOnTableColumn(imageSubtitle, tableIndex, _INSERT_COLUMN_INDEX);
                }

                DocClass.SaveDocAs(this._documentFilePath);
            }
        }
    }
}
