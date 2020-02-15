using System;
using System.IO;
using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using BluePrismInterface.Implementations;
using OfficeExtension;
using Microsoft.Office.Interop.Word;

namespace UnitTests
{
    [TestClass]
    public class WordAdapterUnitTests
    {
        string _ROW_HEADER_LABEL = "Row";
        string _COLUMN_HEADER_LABEL = "Column";
        string _TEXT_HEADER_LABEL = "Text";
        string _TABLE_INDEX_HEADER_LABEL = "TableIndex";
        string TABLE_TITLE = "BreakDown";
        string NEW_ROW_TEXT_CONTENT = "\r\a";
        string DOCUMENT_TEMPLATE_PATH = Directory.GetCurrentDirectory() + @"\Mocks\template.docx";
        string DOCUMENT_FORM_PATH = Directory.GetCurrentDirectory() + @"\Mocks\form.docx";

        [TestMethod]
        public void InserTextIntoWordTableFromDataTable_Success_1()
        {
            //ARRANGE
            System.Data.DataTable dataTable = new System.Data.DataTable();

            dataTable.Columns.Add(_TEXT_HEADER_LABEL, typeof(string));
            dataTable.Columns.Add(_ROW_HEADER_LABEL, typeof(Decimal));
            dataTable.Columns.Add(_COLUMN_HEADER_LABEL, typeof(Decimal));
            dataTable.Columns.Add(_TABLE_INDEX_HEADER_LABEL, typeof(Decimal));

            Random rnd = new Random();

            DataRow row;

            for (int i = 1; i < 6; i++)
            {
                row = dataTable.NewRow();

                row[_TEXT_HEADER_LABEL] = "txt" + i.ToString();
                row[_ROW_HEADER_LABEL] = Convert.ToDecimal(i);
                row[_COLUMN_HEADER_LABEL] = Convert.ToDecimal(rnd.Next(1, 3));
                row[_TABLE_INDEX_HEADER_LABEL] = Convert.ToDecimal(1);

                dataTable.Rows.Add(row);
            }

            //ACT
            int table_row_count_before_insert;
            using (WordDocument DocClass = new WordDocument())
            {
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table_before_inserts = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                table_row_count_before_insert = table_before_inserts.Rows.Count;
            }

            File.Copy(DOCUMENT_TEMPLATE_PATH, DOCUMENT_FORM_PATH, true);
            WordAdapter WordAdap = new WordAdapter(DOCUMENT_FORM_PATH);
            WordAdap.InserTextIntoWordTableFromDataTable(dataTable);


            Table table;
            using (WordDocument DocClass = new WordDocument())
            {
                DocClass.OpenFile(DOCUMENT_FORM_PATH);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);

                //ASSERT
                Assert.IsTrue(table_row_count_before_insert == 1);
                Assert.IsTrue(table.Rows.Count > table_row_count_before_insert);
                foreach (DataRow current_row in dataTable.Rows)
                {
                    string text = current_row[_TEXT_HEADER_LABEL].ToString();
                    int row_index = Decimal.ToInt32((Decimal)current_row[_ROW_HEADER_LABEL]);
                    int colum_index = Decimal.ToInt32((Decimal)current_row[_COLUMN_HEADER_LABEL]);

                    Assert.IsTrue(table.Cell(row_index, colum_index).Range.Text == text + NEW_ROW_TEXT_CONTENT);
                }
            }
        }
    }
}
