using System.Data;

namespace BluePrismInterface.Interfaces
{
    public interface IBluePrismAdapter
    {
        void InsertImagesIntoWordFromDataTable(DataTable datatable);
        void InserTextIntoWordTableFromDataTable(DataTable datatable);
    }
}
