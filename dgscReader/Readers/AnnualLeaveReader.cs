using System;
using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Employees;
using OsanScheduler.Models.Employees.LeaveInfo;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace OsanScheduler.DgscReader.Readers
{
	public class AnnualLeaveReader
	{
		private readonly Site _dgsc;
		private readonly string _file;

		public AnnualLeaveReader(Site dgsc, string file)
		{
			this._dgsc = dgsc;
			this._file = file;
		}

		public Site Process()
		{
			XSSFWorkbook workbook;

			using (FileStream file = new FileStream(this._file,
				FileMode.Open, FileAccess.Read))
			{
				workbook = new XSSFWorkbook(file);
			}

			ISheet sheet = workbook.GetSheetAt(0);

			// the first row of the sheet provides the labels for the data
			// so, create a dictionary of the labels to the cell id or column
			// number to allow for easier recall later.
			var row = sheet.GetRow(0);
			var labels = new Dictionary<string, int>();

			for (int i = 0; i < row.Cells.Count; i++)
			{
				labels[row.GetCell(i).StringCellValue] = i;
			}

			// now process through the rest of the rows getting the leave
			// balance information.  I will only process information for previous
			// and current year.
			var previousyear = DateTime.UtcNow.Year - 1;
			for (int i = 1; i <= sheet.LastRowNum; i++)
			{
				// get the information
				row = sheet.GetRow(i);
				int year = (int)row.GetCell(labels["hYear"]).NumericCellValue;
				var empId = row.GetCell(labels["EmployeeID"]).StringCellValue;
				decimal annual = (decimal)row.GetCell(labels["Annual"]).NumericCellValue;
				decimal carry = (decimal)row.GetCell(labels["CarryOver"]).NumericCellValue;
				if (year >= previousyear)
				{
					int empPos = -1;
					for (int e = 0; e < this._dgsc.Employees.Count && empPos < 0; e++)
					{
						Employee emp = this._dgsc.Employees[e];
						if (emp.CompanyInfo.CompanyID.Equals(empId))
						{
							empPos = e;
							var balPos = -1;
							for (var b = 0; b < emp.Balances.Count && balPos < 0; b++)
							{
								var bal = emp.Balances[b];
								if (bal.Year == year)
								{
									bal.AnnualLeave = annual;
									bal.CarryOver = carry;
									emp.Balances[b] = bal;
									balPos = b;
								}
							}
							if (balPos < 0)
							{
								var bal = new LeaveBalance();
								bal.Year = year;
								bal.AnnualLeave = annual;
								bal.CarryOver = carry;
								emp.Balances.Add(bal);
							}
							this._dgsc.Employees[empPos] = emp;
						}
					}
				}
			}

			return this._dgsc;
		}
	}
}

