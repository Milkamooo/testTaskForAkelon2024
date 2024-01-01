using ClosedXML.Excel;

namespace task4.Services
{
	public class FileUpdateService
	{
        public bool UpdateExcel(string filePath, string organizationName, string newContactName)
        {
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(2);
            var row = worksheet.Search(organizationName).FirstOrDefault().Address.RowNumber;
            if (row != null)
            {
                worksheet.Cell(row, 4).Value = newContactName;
                workbook.Save();
                return true;
            }
            else return false;
        }
	}
}