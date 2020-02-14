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
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("Text");
            dataTable.Columns.Add("Row", typeof(Int32));
            dataTable.Columns.Add("Column", typeof(Int32));
            dataTable.Columns.Add("TableIndex", typeof(Int32));

            Random rnd = new Random();

            DataRow row = null;

            for (int i = 1; i < 5; i++)
            {
                row = dataTable.NewRow();

                row.SetField("Text", "txt" + i.ToString());
                row.SetField("Row", i);
                row.SetField("Column", rnd.Next(1,3));
                row.SetField<int>("TableIndex", 1);

                dataTable.Rows.Add(row);
            }


            File.Copy(@"C:\Users\p.de.barros.mesquita\source\repos\ARMS_Integracao\OfficeExtension\DocImage\template.docx", @"C:\Users\p.de.barros.mesquita\source\repos\ARMS_Integracao\OfficeExtension\DocImage\form.docx", true);
            WordAdapter WordAdap = new WordAdapter(@"C:\Users\p.de.barros.mesquita\source\repos\ARMS_Integracao\OfficeExtension\DocImage\form.docx");
            WordAdap.InserTextIntoWordTableFromDataTable(dataTable);

        }
    }
}
