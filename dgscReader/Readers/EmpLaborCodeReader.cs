using System;
using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Sites.Labor;
using OsanScheduler.Models.Employees;
using OsanScheduler.Models.Employees.Labor;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace OsanScheduler.DgscReader.Readers
{
	public class EmpLaborCodeReader
	{
		private readonly Site _dgsc;
		private readonly string _file;

		public EmpLaborCodeReader(Site dgsc, string file)
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
				int year = (int)row.GetCell(labels["FiscalYear"]).NumericCellValue;

				var empId = "";
				var cell = row.GetCell(labels["EmployeeID"]);
				if (cell != null)
                {
					empId = cell.StringCellValue;
                }
				if (!empId.Equals(""))
                {
					var chgNo = row.GetCell(labels["WorkCode"]).StringCellValue;
					var ext = row.GetCell(labels["Extension"]).StringCellValue;
					if (year >= previousyear)
					{
						SiteLaborCode? siteLabor = null;
						this._dgsc.LaborCodes.ForEach(slc =>
						{
							if (slc.ChargeNumber.Equals(chgNo)
								&& slc.Extension.Equals(ext))
							{
								siteLabor = slc;
							}
						});
						var empPos = -1;
						for (int e = 0; e < this._dgsc.Employees.Count && empPos < 0; e++)
						{
							var emp = this._dgsc.Employees[e];
							if (emp.CompanyInfo.EmployeeID.Equals(empId))
							{
								empPos = e;
								var found = false;
								for (int l = 0; l < emp.Labor.Count && !found; l++)
								{
									var lc = emp.Labor[l];
									if (lc.ChargeNumber.Equals(chgNo)
										&& lc.Extension.Equals(ext))
									{
										found = true;
									}
								}
								if (!found)
								{
									var lc = new EmployeeLaborCode();
									lc.ChargeNumber = chgNo;
									lc.Extension = ext;
									lc.CompanyCode = emp.CompanyInfo.CompanyCode;
									if (siteLabor != null)
									{
										lc.StartDate = new DateTime(siteLabor.StartDate.Ticks);
										if (siteLabor.EndDate < emp.Assignments[
											emp.Assignments.Count - 1].EndDate)
										{
											lc.EndDate = new DateTime(siteLabor.EndDate.Ticks);
										}
										else
										{
											lc.EndDate = new DateTime(emp.Assignments[
												emp.Assignments.Count - 1].EndDate.Ticks);
										}
									}
									emp.Labor.Add(lc);
								}
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

