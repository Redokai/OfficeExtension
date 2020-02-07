using OfficeExtension;

namespace DocImage
{
    class Program
    {
        static void Main(string[] args)
        {
            string TABLE_TITLE = "BreakDown";
            int COLUMN_INDEX = 0;
            string IMAGE_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\gaivota.jpg";
            string DOCUMENT_TEMPLATE_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\template.docx";
            string DOCUMENT_OUTPUT_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\form.docx";

            WordDocument DocClass = new WordDocument();
            DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
            DocClass.AppendImageOnTableColumn(IMAGE_PATH, TABLE_TITLE, COLUMN_INDEX);
            DocClass.SaveDocAs(DOCUMENT_OUTPUT_PATH);

        }
    }
}
