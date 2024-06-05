using Excel_Modification_Task.Models;
using OfficeOpenXml;

namespace Excel_Modification_Task.Helpers
{
	public static class DocumentSetting
	{
		public static FileData ProcessFile(string filePath)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			var Data = new List<List<string>>();
			double grandTotalValueAfterTax = 0;
			using(var package = new ExcelPackage(new FileInfo(filePath)))
			{
				var excelSheet = package.Workbook.Worksheets[0];
				var rowsCount = excelSheet.Dimension.Rows;
				var columnsCount = excelSheet.Dimension.Columns;

				//adding new header cell
				excelSheet.Cells[1, columnsCount + 1].Value = "Total Value before Taxing";
                
				
				var firstRow = new List<string>();
				for (global::System.Int32 j = 1; j <= columnsCount+1; j++)
                {
					firstRow.Add(excelSheet.Cells[1, j].GetValue<string>());
                }

                Data.Add(firstRow);
                for (int i = 2; i <= rowsCount; i++)
                {
					double totalValueAfterTax = excelSheet.Cells[i, 7].GetValue<double>();
					double taxValue = excelSheet.Cells[i, 8].GetValue<double>();
					excelSheet.Cells[i, columnsCount + 1].Value =Math.Round( (totalValueAfterTax - taxValue),2);
					grandTotalValueAfterTax += excelSheet.Cells[i, 7].GetValue<double>();

					var row = new List<string>();
                    for (int j = 1; j <= columnsCount+1; j++)
                    {
						row.Add(excelSheet.Cells[i,j].GetValue<string>());
                    }
					Data.Add(row);
                }
				package.Save();
            }
			return new FileData()
			{
				data = Data,
				Sum = grandTotalValueAfterTax,
				filePath=filePath
			};
		}
	}
}
