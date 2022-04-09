using System;
using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Sites.Labor;
using OsanScheduler.Models.Employees;
using OsanScheduler.Models.Employees.Labor;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace OsanScheduler.DgscReader.Readers
{
	public class SiteLaborCodeReader
	{
		private readonly Site _dgsc;
		private readonly string _file;

		public SiteLaborCodeReader(Site dgsc, string file)
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
				var chgNo = row.GetCell(labels["WorkCode"]).StringCellValue;
				var ext = row.GetCell(labels["Extension"]).StringCellValue;
				var clin = row.GetCell(labels["CLIN"]).StringCellValue;
				var slin = row.GetCell(labels["SLIN"]).StringCellValue;
				var loc = row.GetCell(labels["Location"]).StringCellValue;
				var wbs = row.GetCell(labels["WBS"]).StringCellValue;
				var min = (int) row.GetCell(labels["MinimumEmployees"]).NumericCellValue;
				var noname = row.GetCell(labels["NoEmployeeName"]).StringCellValue;
				var hours = (decimal)row.GetCell(labels["HoursPerEmployee"]).NumericCellValue;
				var bEx = row.GetCell(labels["ExerciseCode"]).BooleanCellValue;
				var startdate = row.GetCell(labels["StartDate"]).DateCellValue;
				var enddate = row.GetCell(labels["EndDate"]).DateCellValue;
				if (year >= previousyear)
                {
					var found = false;
					for (int l = 0; l < this._dgsc.LaborCodes.Count && !found; l++)
                    {
						var lbr = this._dgsc.LaborCodes[l];
						if (lbr.Equals("raytheon", chgNo, ext))
                        {
							found = true;
							lbr.Clin = clin;
							lbr.Slin = slin;
							lbr.Location = loc;
							lbr.Wbs = wbs;
							lbr.Minimum = min;
							lbr.NoEmployee = noname;
							lbr.ContractHours = hours;
							lbr.IsExercise = bEx;
							lbr.StartDate = startdate;
							lbr.EndDate = enddate;
							this._dgsc.LaborCodes[l] = lbr;
                        }
                    }
					if (!found)
                    {
						var lbr = new SiteLaborCode();
						lbr.Company = "raytheon";
						lbr.ChargeNumber = chgNo;
						lbr.Extension = ext;
						lbr.Division = "RTSC";
						lbr.Clin = clin;
						lbr.Slin = slin;
						lbr.Location = loc;
						lbr.Wbs = wbs;
						lbr.Minimum = min;
						lbr.NoEmployee = noname;
						lbr.ContractHours = hours;
						lbr.IsExercise = bEx;
						lbr.StartDate = startdate;
						lbr.EndDate = enddate;
						this._dgsc.LaborCodes.Add(lbr);
					}
                }
			}
			return this._dgsc;
		}
	}
}

