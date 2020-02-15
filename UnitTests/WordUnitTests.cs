using OfficeExtension;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Office.Interop.Word;
using System;
using System.IO;

namespace UnitTests
{
    [TestClass]
    public class FindTableUnitTests
    {
        string TABLE_TITLE_SUCCESS = "BreakDown";
        string TABLE_TITLE_FAILURE = "breakdown";
        string DOCUMENT_TEMPLATE_PATH = Directory.GetCurrentDirectory() + @"\Mocks\template.docx";

        [TestMethod]
        public void Find_Success_1()
        {
            //LOOKS FOR THE FIRST EMPTY CELL ON A DETERMINED TABLE, EXPECTED RESULT: FIND TABLE & FIND EMPTY CELL

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE_SUCCESS };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);

                //ASSERT
                Assert.IsNotNull(table);
            }
        }

        [TestMethod]
        public void Find_Failure_1()
        {
            //LOOKS FOR THE FIRST EMPTY CELL ON A DETERMINED TABLE, EXPECTED RESULT: DO NOT FIND TABLE NOR FIND EMPTY CELL

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE_FAILURE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);

                //ASSERT
                Assert.IsNull(table);
            }
        }
    }

    [TestClass]
    public class FindEmptyCellsUnitTests
    {
        string TABLE_TITLE = "BreakDown";
        int COLUMN_INDEX = 0;
        string NEW_ROW_TEXT_CONTENT = "\r\a";
        string DOCUMENT_TEMPLATE_PATH_SUCCESS = Directory.GetCurrentDirectory() + @"\Mocks\template.docx";
        string DOCUMENT_TEMPLATE_PATH_FAILURE = Directory.GetCurrentDirectory() + @"\Mocks\form.docx";

        [TestMethod]
        public void Find_Success_1()
        {
            //CHECK IF FIRST EMPTY CELL FOUND IS CURRENTLY EMPTY, EXPECTED: EMPTY

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH_SUCCESS);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                object[] find_empty_cells_args = new object[2] { table, COLUMN_INDEX };
                Range first_empty_cell_on_template = (Range)DocPriv.Invoke("_FindTableEmptyCellOnSpecificCollumn", find_empty_cells_args);

                //ASSERT
                Assert.IsTrue(first_empty_cell_on_template.InlineShapes.Count == 0);
                Assert.IsTrue(first_empty_cell_on_template.Text == NEW_ROW_TEXT_CONTENT);
            }
        }

        public void Find_Failure_1()
        {
            //CHECK IF THERE IS AN EMPTY CELL ON TABLE COLUMN, EXPECTED: EMPTY CELL NOT FOUND, NULL VALUE

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH_FAILURE);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                object[] find_empty_cells_args = new object[2] { table, COLUMN_INDEX };
                Range first_empty_cell_on_template = (Range)DocPriv.Invoke("_FindTableEmptyCellOnSpecificCollumn", find_empty_cells_args);

                //ASSERT
                Assert.IsNull(first_empty_cell_on_template.Text);
                Assert.IsTrue(first_empty_cell_on_template.InlineShapes.Count == 0);
            }
        }

        public void Find_Failure_2()
        {
            //LOOKS FOR THE FIRST EMPTY CELL ON A DETERMINED TABLE 
            // AFTER SETTING THE ONLY CELL TO NON NULL VALUE
            //, EXPECTED RESULT: FIND TABLE BUT NOT FIND EMPTY CELL

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH_FAILURE);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                object[] find_empty_cells_args = new object[2] { table, COLUMN_INDEX };
                Range first_try_empty_cell_on_template = (Range)DocPriv.Invoke("_FindTableEmptyCellOnSpecificCollumn", find_empty_cells_args);
                first_try_empty_cell_on_template.Text = "TESTE";
                Range second_try_empty_cell_on_template = (Range)DocPriv.Invoke("_FindTableEmptyCellOnSpecificCollumn", find_empty_cells_args);

                //ASSERT
                Assert.IsNotNull(first_try_empty_cell_on_template);
                Assert.IsNull(second_try_empty_cell_on_template.InlineShapes.Count == 0);
            }
        }
    }

    [TestClass]
    public class AppendTableRowUnitTests
    {
        string TABLE_TITLE = "BreakDown";
        string TABLE_TITLE_FAILURE = "breakdown";
        string NEW_ROW_TEXT_CONTENT = "\r\a";
        string TEST_TEXT = "TESTE";
        int COLUMN_INDEX = 1;
        int ROW_INDEX = 1;
        string DOCUMENT_TEMPLATE_PATH_SUCCESS = Directory.GetCurrentDirectory() + @"\Mocks\form.docx";

        [TestMethod]
        public void Append_Success_1()
        {
            //CHECKS IF ROW WAS SUCCESSFULLY APPENDED IN THE TABLE

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH_SUCCESS);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                int tableRowCountBeforeAppend = (int)table.Rows.Count;
                object[] table_as_args = new object[1] { table };
                DocPriv.Invoke("_AppendTableRow", table_as_args);
                int tableRowCountAfterAppend = (int)table.Rows.Count;

                //ASSERT
                Assert.IsTrue(tableRowCountAfterAppend > tableRowCountBeforeAppend);
            }
        }

        [TestMethod]
        public void Append_Success_2()
        {
            //CHECKS IF ROW WAS SUCCESSFULLY APPENDED AFTER DETERMINED ROW

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH_SUCCESS);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                Cell first_cell = table.Cell(ROW_INDEX, COLUMN_INDEX);
                first_cell.Range.Text = TEST_TEXT;
                int tableRowCountBeforeAppend = (int)table.Rows.Count;
                object[] table_as_args = new object[1] { table };
                DocPriv.Invoke("_AppendTableRow", table_as_args);
                int tableRowCountAfterAppend = (int)table.Rows.Count;
                first_cell = table.Cell(ROW_INDEX, COLUMN_INDEX);
                Cell second_cell = table.Cell(ROW_INDEX + 1, COLUMN_INDEX);
                string first_cell_text = first_cell.Range.Text;
                string second_cell_text = second_cell.Range.Text;

                //ASSERT
                Assert.IsTrue(first_cell_text == TEST_TEXT + NEW_ROW_TEXT_CONTENT);
                Assert.IsTrue(second_cell_text == NEW_ROW_TEXT_CONTENT);
                Assert.IsFalse(second_cell_text == TEST_TEXT + NEW_ROW_TEXT_CONTENT);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(NullReferenceException))]
        public void Append_Failure_1()
        {
            //ASSERTS ERROR IF NO TABLE IS FOUND

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH_SUCCESS);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE_FAILURE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                int tableRowCountBeforeAppend = (int)table.Rows.Count;
                object[] table_as_args = new object[1] { table };
                DocPriv.Invoke("_AppendTableRow", table_as_args);
                int tableRowCountAfterAppend = (int)table.Rows.Count;
            }
        }
    }

    [TestClass]
    public class AppendImageUnitTests
    {
        string TABLE_TITLE = "BreakDown";
        int COLUMN_INDEX = 1;
        int ROW_INDEX = 1;
        string IMAGE_PATH = Directory.GetCurrentDirectory() + @"\Mocks\gaivota.jpg";
        string DOCUMENT_TEMPLATE_PATH = Directory.GetCurrentDirectory() + @"\Mocks\template.docx";

        [TestMethod]
        public void Insert_Success_1()
        {
            // ASSERTS IF IMAGE IS INSERTED ON PARTICULAR CELL

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                Cell first_cell = table.Cell(ROW_INDEX, COLUMN_INDEX);
                int first_cell_img_count_before = first_cell.Range.InlineShapes.Count;
                DocClass.AppendImageOnTableColumn(IMAGE_PATH, TABLE_TITLE, COLUMN_INDEX);
                int first_cell_img_count_after = first_cell.Range.InlineShapes.Count;

                //ASSERT
                Assert.IsTrue(first_cell_img_count_before == 0);
                Assert.IsTrue(first_cell_img_count_after == 1);
            }
        }
    }

    [TestClass]
    public class AppendTestUnitTests
    {
        string TABLE_TITLE = "BreakDown";
        int COLUMN_INDEX = 1;
        int ROW_INDEX = 1;
        string NEW_ROW_TEXT_CONTENT = "\r\a";
        string TEST_TEXT = "TESTE";
        string DOCUMENT_TEMPLATE_PATH = Directory.GetCurrentDirectory() + @"\Mocks\template.docx";

        [TestMethod]
        public void Insert_Success_1()
        {
            // ASSERTS IF IMAGE IS INSERTED ON PARTICULAR CELL

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                Cell first_cell = table.Cell(ROW_INDEX, COLUMN_INDEX);
                string first_cell_text_before = first_cell.Range.Text;
                DocClass.AppendTextOnTableColumn(TEST_TEXT, TABLE_TITLE, COLUMN_INDEX);
                string first_cell_text_after = first_cell.Range.Text;

                //ASSERT
                Assert.IsTrue(first_cell_text_before == NEW_ROW_TEXT_CONTENT);
                Assert.IsTrue(first_cell_text_after == TEST_TEXT + NEW_ROW_TEXT_CONTENT);
            }
        }
    }

    [TestClass]
    public class _InsertTextInRangeUnitTests
    {
        string TABLE_TITLE = "BreakDown";
        string TEST_TEXT = "test";
        string NEW_ROW_TEXT_CONTENT = "\r\a";
        int COLUMN_INDEX_SUCCESS = 1;
        int ROW_INDEX_SUCCESS = 1;
        int COLUMN_INDEX_FAILURE = 3;
        int ROW_INDEX_FAILURE = 7;
        string DOCUMENT_TEMPLATE_PATH = Directory.GetCurrentDirectory() + @"\Mocks\template.docx";

        [TestMethod]
        public void _InsertTextInRange_Success_1()
        {
            // ASSERTS IF CELL TEXT IS CHANGED AFTER INSERT

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                Range first_cell_range = table.Cell(ROW_INDEX_SUCCESS, COLUMN_INDEX_SUCCESS).Range;
                string first_cell_text_before_insert = first_cell_range.Text;
                object[] insert_text_args = new object[2] { TEST_TEXT, first_cell_range };
                DocPriv.Invoke("_InsertTextInRange", insert_text_args);
                string first_cell_text_after_insert = first_cell_range.Text;

                //ASSERT
                Assert.IsTrue(first_cell_text_before_insert == NEW_ROW_TEXT_CONTENT);
                Assert.IsTrue(first_cell_text_after_insert == TEST_TEXT + NEW_ROW_TEXT_CONTENT);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(System.Runtime.InteropServices.COMException))]
        public void _InsertTextInRange_Failure_1()
        {
            // ASSERTS IF NON EXISTENT CELLS OPERATION LEADS TO System.Runtime.InteropServices.COMException

            //ARRANGE
            using (WordDocument DocClass = new WordDocument())
            {
                //ACT
                DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
                PrivateObject DocPriv = new PrivateObject(DocClass);
                object[] table_title_as_args = new object[1] { TABLE_TITLE };
                Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
                Range first_cell_range = table.Cell(ROW_INDEX_FAILURE, COLUMN_INDEX_FAILURE).Range;
                string first_cell_text_before_insert = first_cell_range.Text;
                object[] insert_text_args = new object[2] { TEST_TEXT, first_cell_range };
                DocPriv.Invoke("_InsertTextInRange", insert_text_args);
                string first_cell_text_after_insert = first_cell_range.Text;
            }
        }
    }
}
