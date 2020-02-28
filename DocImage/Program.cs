using System;
using System.Data;
using System.IO;
using BluePrismInterface.Implementations;

namespace DocImage
{
    class Program
    {
        static void Main(string[] args)
        {
            string DOCUMENT_TEMPLATE_PATH = Directory.GetCurrentDirectory() + @"\template.docx";
            string DOCUMENT_FORM_PATH = Directory.GetCurrentDirectory() + @"\form.docx";

            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("Text");
            dataTable.Columns.Add("Row", typeof(Decimal));
            dataTable.Columns.Add("Column", typeof(Decimal));
            dataTable.Columns.Add("TableIndex", typeof(Decimal));

            Random rnd = new Random();

            DataRow row = null;

            for (int i = 1; i < 2; i++)
            {
                row = dataTable.NewRow();

                row.SetField("Text", "Replaced Successfully");
                row.SetField("Token", "$FindAndReplaceMe$");

                dataTable.Rows.Add(row);
            }
            try
            {

            }
            catch (Exception err)
            {

                throw(err);
            }

            File.Copy(DOCUMENT_TEMPLATE_PATH, DOCUMENT_FORM_PATH, true);
            WordAdapter WordAdap = new WordAdapter(DOCUMENT_FORM_PATH);
            WordAdap.ReplaceTexts(dataTable);

        }
    }
}
