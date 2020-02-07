using OfficeExtension;
using System;

namespace DocImage
{
    class Program
    {
        static void Main(string[] args)
        {
            string TABLE_TITLE = "BreakDown";
            int COLUMN_INDEX = 1;
            int ROW_INDEX = 1;
            string IMAGE_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\gaivota.jpg";
            string DOCUMENT_TEMPLATE_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\template.docx";
            string DOCUMENT_OUTPUT_PATH = @"C:\Users\Red\source\repos\DocImage\DocImage\form.docx";

            ;

            using (WordDocument DocClass = new WordDocument())
            {
                try
                {
                    DocClass.OpenFile(DOCUMENT_TEMPLATE_PATH);
                    DocClass.AppendImageOnTableColumn(IMAGE_PATH, TABLE_TITLE, COLUMN_INDEX);
                    DocClass.AppendTextOnTableColumn("TESTE", TABLE_TITLE, COLUMN_INDEX);
                    DocClass.AppendImageOnTableColumn(IMAGE_PATH, TABLE_TITLE, COLUMN_INDEX);
                    DocClass.AppendTextOnTableColumn("TESTE2", TABLE_TITLE, COLUMN_INDEX);
                    DocClass.AppendImageOnTableColumn(IMAGE_PATH, TABLE_TITLE, COLUMN_INDEX);
                    DocClass.AppendTextOnTableColumn("TESTE3", TABLE_TITLE, COLUMN_INDEX);
                    DocClass.SaveDocAs(DOCUMENT_OUTPUT_PATH);
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine(ex.StackTrace);
                    Console.ReadKey();
                }
            }
        }
    }
}
