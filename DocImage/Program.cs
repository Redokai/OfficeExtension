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

            dataTable.Columns.Add("FileImage");
            dataTable.Columns.Add("Texto");
            dataTable.Columns.Add("TableIndex", typeof(Int32));

            DataRow row = null;

            for (int i = 1; i < 2; i++)
            {
                row = dataTable.NewRow();

                row.SetField("FileImage", @"C:\Users\p.de.barros.mesquita\source\repos\ARMS_Integracao\OfficeExtension\DocImage\gaivota.jpg");
                row.SetField("Texto", $"Aqui vai um texto da tabela 1 {i}!");
                row.SetField<int>("TableIndex", 1);

                dataTable.Rows.Add(row);
            }

            for (int i = 1; i < 10; i++)
            {
                row = dataTable.NewRow();

                row.SetField("FileImage", @"C:\Users\p.de.barros.mesquita\source\repos\ARMS_Integracao\OfficeExtension\DocImage\gaivota.jpg");
                row.SetField("Texto", $"Aqui vai um texto da tabela 2 {i}!");
                row.SetField<int>("TableIndex", 2);

                dataTable.Rows.Add(row);
            }

            File.Copy(@"C:\Users\p.de.barros.mesquita\source\repos\ARMS_Integracao\OfficeExtension\DocImage\template.docx", @"C:\Users\p.de.barros.mesquita\source\repos\ARMS_Integracao\OfficeExtension\DocImage\form.docx", true);
            WordAdapter WordAdap = new WordAdapter(@"C:\Users\p.de.barros.mesquita\source\repos\ARMS_Integracao\OfficeExtension\DocImage\form.docx");
            WordAdap.InsertImagesIntoWordFromDataTable(dataTable);

        }
    }
}
