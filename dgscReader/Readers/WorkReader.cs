using System;
using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Employees.Labor;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace OsanScheduler.DgscReader.Readers
{
	public class WorkReader
	{
		private readonly Site _dgsc;
		private readonly string _file;

		public WorkReader(Site dgsc, string file)
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

			// now process through the rest of the rows to get all the employees
			// that need to be processed.
			// we will limit the employees to those who haven't left or have
			// only left within the last year.
			int previousyear = DateTime.UtcNow.Year - 1;

			DateTime baseDate = new DateTime(1970, 1, 1);
			DateTime expireDate = new DateTime(previousyear, 1, 1);

			for (int i = 1; i <= sheet.LastRowNum; i++)
			{
				row = sheet.GetRow(i);
				var empID = row.GetCell(labels["EmployeeID"]).StringCellValue;
				var date = row.GetCell(labels["DateWorked"]).DateCellValue;
				if (date >= expireDate)
				{
					var chgNo = row.GetCell(labels["ChargeNumber"]).StringCellValue;
					var ext = "";
					if (row.GetCell(labels["Extension"]) != null)
					{
						ext = row.GetCell(labels["Extension"]).StringCellValue;
					}
					var hours = (decimal)row.GetCell(labels["Hours"]).NumericCellValue;

					var found = false;
					for (int e = 0; e < this._dgsc.Employees.Count && !found; e++)
					{
						var emp = this._dgsc.Employees[e];
						if (emp.CompanyInfo.EmployeeID.Equals(empID))
						{
							for (int w = 0; w < emp.Work.Count && !found; w++)
							{
								var work = emp.Work[w];
								if (work.Equals(date, emp.CompanyInfo.CompanyCode, chgNo, ext))
								{
									found = true;
									work.Hours = hours;
									emp.Work[w] = work;
								}
							}
							if (!found)
							{
								var work = new Work();
								work.EmployeeId = emp.Id;
								work.DateWorked = date;
								work.CompanyId = emp.CompanyInfo.CompanyCode;
								work.ChargeNumber = chgNo;
								work.Extension = ext;
								work.Hours = hours;
								emp.Work.Add(work);
							}
							this._dgsc.Employees[e] = emp;
						}
					}
				}
			}
			return this._dgsc;
		}
	}
}

