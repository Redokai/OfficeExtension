using OfficeExtension;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Office.Interop.Word;

namespace UnitTests
{
    [TestClass]
    public class FindTableUnitTests
    {
        string TABLE_TITLE_SUCCESS = "BreakDown";
        string TABLE_TITLE_FAILURE = "breakdown";
        int COLUMN_INDEX = 0;
        string IMAGE_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\gaivota.jpg";
        string DOCUMENT_TEMPLATE_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\template.docx";
        string DOCUMENT_OUTPUT_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\form.docx";

        [TestMethod]
        public void Find_Success_1()
        {
            //ARRANGE
            WordDocument DocClass = new WordDocument();

            //ACT
            DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
            PrivateObject DocPriv = new PrivateObject(DocClass);
            object[] table_title_as_args = new object[1] { TABLE_TITLE_SUCCESS };
            Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
            DocClass.Close();
            DocClass.Quit();

            //ASSERT
            Assert.IsNotNull(table);
        }

        [TestMethod]
        public void Find_Failure_1()
        {
            //ARRANGE
            WordDocument DocClass = new WordDocument();

            //ACT
            DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
            PrivateObject DocPriv = new PrivateObject(DocClass);
            object[] table_title_as_args = new object[1] { TABLE_TITLE_FAILURE };
            Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
            DocClass.Close();
            DocClass.Quit();

            //ASSERT
            Assert.IsNull(table);
        }
    }

    [TestClass]
    public class FindEmptyCellsUniteTests
    {
        string TABLE_TITLE = "BreakDown";
        int COLUMN_INDEX = 0;
        string IMAGE_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\gaivota.jpg";
        string DOCUMENT_TEMPLATE_PATH_SUCCESS = @"C:\Users\Red\source\repos\DocImage\UnitTests\Mocks\form.docx";
        string DOCUMENT_TEMPLATE_PATH_FAILURE = @"C:\Users\Red\source\repos\DocImage\UnitTests\Mocks\form.docx";

        [TestMethod]
        public void Find_Success_1()
        {
            //CHECK IF FIRST EMPTY CELL FOUND IS CURRENTLY EMPTY, EXPECTED: EMPTY

            //ARRANGE
            WordDocument DocClass = new WordDocument();

            //ACT
            DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH_SUCCESS);
            PrivateObject DocPriv = new PrivateObject(DocClass);
            object[] table_title_as_args = new object[1] { TABLE_TITLE };
            Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
            object[] find_empty_cells_args = new object[2] { table, COLUMN_INDEX };
            Range first_empty_cell_on_template = (Range)DocPriv.Invoke("_FindTableEmptyCellOnSpecificCollumn", find_empty_cells_args);
            DocClass.Close();
            DocClass.Quit();

            //ASSERT
            Assert.IsNull(first_empty_cell_on_template.Text);
            Assert.IsTrue(first_empty_cell_on_template.InlineShapes.Count == 0);
        }

        public void Find_Failure_1()
        {
            //CHECK IF THERE IS AN EMPTY CELL ON TABLE COLUMN, EXPECTED: EMPTY CELL NOT FOUND, NULL VALUE

            //ARRANGE
            WordDocument DocClass = new WordDocument();

            //ACT
            DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH_FAILURE);
            PrivateObject DocPriv = new PrivateObject(DocClass);
            object[] table_title_as_args = new object[1] { TABLE_TITLE };
            Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
            object[] find_empty_cells_args = new object[2] { table, COLUMN_INDEX };
            Range first_empty_cell_on_template = (Range)DocPriv.Invoke("_FindTableEmptyCellOnSpecificCollumn", find_empty_cells_args);
            DocClass.Close();
            DocClass.Quit();

            //ASSERT
            Assert.IsNull(first_empty_cell_on_template.Text);
            Assert.IsTrue(first_empty_cell_on_template.InlineShapes.Count == 0);
        }
    }

    [TestClass]
    public class ImageInsertUnitTests
    {
        string TABLE_TITLE = "BreakDown";
        int COLUMN_INDEX = 0;
        string IMAGE_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\gaivota.jpg";
        string DOCUMENT_TEMPLATE_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\template.docx";
        string DOCUMENT_OUTPUT_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\form.docx";

        [TestMethod]
        public void Insert_Success_1()
        {
            //ARRANGE
            WordDocument DocClass = new WordDocument();

            //ACT
            DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
            PrivateObject DocPriv = new PrivateObject(DocClass);
            object[] table_title_as_args = new object[1] { TABLE_TITLE };
            Table table = (Table)DocPriv.Invoke("_FindTable", table_title_as_args);
            object[] find_empty_cells_args = new object[2] { table, COLUMN_INDEX };
            Range first_empty_cell_on_template = (Range)DocPriv.Invoke("_FindTableEmptyCellOnSpecificCollumn", find_empty_cells_args);
            int shape_count_before_insert = first_empty_cell_on_template.InlineShapes.Count;
            DocClass.AppendImageOnTableColumn(IMAGE_PATH, TABLE_TITLE, COLUMN_INDEX);
            int shape_count_after_insert = first_empty_cell_on_template.InlineShapes.Count;
            DocClass.Close();
            DocClass.Quit();

            //ASSERT
            Assert.IsTrue(shape_count_before_insert == 0);
            Assert.IsTrue(shape_count_after_insert == 1);
        }
    }
}
