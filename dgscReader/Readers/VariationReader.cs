using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Employees.Assignments;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace OsanScheduler.DgscReader.Readers
{
	public class VariationReader
	{
		private readonly Site _dgsc;
		private readonly string _file;

		public VariationReader(Site dgsc, string file)
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
				var isMids = (row.GetCell(labels["VariationType"]).StringCellValue.Equals("MIDS"));
				var code = row.GetCell(labels["ShowCode"]).StringCellValue;
				var empID = row.GetCell(labels["EmployeeID"]).StringCellValue;
				var startdate = row.GetCell(labels["StartDate"]).DateCellValue;
				var enddate = row.GetCell(labels["EndDate"]).DateCellValue;
				var daysOff = row.GetCell(labels["DaysOff"]).StringCellValue;

				var found = false;
				for (int e=0; e < this._dgsc.Employees.Count && !found; e++)
                {
					var emp = this._dgsc.Employees[e];
					var workcenter = "";
					var starttime = -1;
					var hours = 8M;
					if (daysOff.Length == 3)
					{
						hours = 10M;
					}
					for (int s = 0; s < this._dgsc.WorkCodes.Count
						&& starttime < 0; s++)
					{
						if (this._dgsc.WorkCodes[s].Code.Equals(code))
						{
							starttime = this._dgsc.WorkCodes[s].StartTime;
						}
					}
					if (isMids)
					{
						workcenter = "GEOINT";
					}
					else
					{
						var sch = emp.Assignments[
							emp.Assignments.Count - 1].Schedules[0];
						for (int d = 0; d < sch.Workdays.Count && workcenter.Equals(""); d++)
						{
							if (!sch.Workdays[d].Workcenter.Equals(""))
							{
								workcenter = sch.Workdays[d].Workcenter;
							}
						}
					}
					if (emp.CompanyInfo.EmployeeID.Equals(empID))
                    {
						for (int v=0; v < emp.Variations.Count && !found; v++)
                        {
							var vari = emp.Variations[v];
							if (vari.IsMids == isMids
								&& vari.StartDate.Equals(startdate)
								&& vari.EndDate.Equals(enddate))
                            {
								found = true;
								for (int d=0; d < 7; d++)
                                {
									if ((d == 0 && (!daysOff.Equals("SS")
										&& !daysOff.Equals("SM")
										&& !daysOff.Equals("FSS")
										&& !daysOff.Equals("SSM")
										&& !daysOff.Equals("SMT")))
										|| (d == 1 && (!daysOff.Equals("MT")
										&& !daysOff.Equals("SM")
										&& !daysOff.Equals("MTW")
										&& !daysOff.Equals("SSM")
										&& !daysOff.Equals("SMT")))
										|| (d == 2 && (!daysOff.Equals("MT")
										&& !daysOff.Equals("TW")
										&& !daysOff.Equals("SMT")
										&& !daysOff.Equals("MTW")
										&& !daysOff.Equals("TWT")))
										|| (d == 3 && (!daysOff.Equals("TW")
										&& !daysOff.Equals("WT")
										&& !daysOff.Equals("MTW")
										&& !daysOff.Equals("TWT")
										&& !daysOff.Equals("WTF")))
										|| (d == 4 && (!daysOff.Equals("WT")
										&& !daysOff.Equals("TF")
										&& !daysOff.Equals("TWT")
										&& !daysOff.Equals("WTF")
										&& !daysOff.Equals("TFS")))
										|| (d == 5 && (!daysOff.Equals("TF")
										&& !daysOff.Equals("FS")
										&& !daysOff.Equals("WTF")
										&& !daysOff.Equals("TFS")
										&& !daysOff.Equals("FSS")))
										|| (d == 6 && (!daysOff.Equals("FS")
										&& !daysOff.Equals("SS")
										&& !daysOff.Equals("TFS")
										&& !daysOff.Equals("FSS")
										&& !daysOff.Equals("SSM")))
										)
                                    {
										vari.SetWorkday(d, workcenter, starttime,
											code, hours);
                                    } else
                                    {
										vari.SetWorkday(d, "", 0, "", 0M);
                                    }
                                }
								emp.Variations[v] = vari;
                            }
                        }
						if (!found)
						{
							var vari = new Variation();
							vari.EndDate = enddate;
							vari.StartDate = startdate;
							vari.IsMids = isMids;
							vari.Site = this._dgsc.Code;

							for (int d = 0; d < 7; d++)
							{
								if ((d == 0 && (!daysOff.Equals("SS")
									&& !daysOff.Equals("SM")
									&& !daysOff.Equals("FSS")
									&& !daysOff.Equals("SSM")
									&& !daysOff.Equals("SMT")))
									|| (d == 1 && (!daysOff.Equals("MT")
									&& !daysOff.Equals("SM")
									&& !daysOff.Equals("MTW")
									&& !daysOff.Equals("SSM")
									&& !daysOff.Equals("SMT")))
									|| (d == 2 && (!daysOff.Equals("MT")
									&& !daysOff.Equals("TW")
									&& !daysOff.Equals("SMT")
									&& !daysOff.Equals("MTW")
									&& !daysOff.Equals("TWT")))
									|| (d == 3 && (!daysOff.Equals("TW")
									&& !daysOff.Equals("WT")
									&& !daysOff.Equals("MTW")
									&& !daysOff.Equals("TWT")
									&& !daysOff.Equals("WTF")))
									|| (d == 4 && (!daysOff.Equals("WT")
									&& !daysOff.Equals("TF")
									&& !daysOff.Equals("TWT")
									&& !daysOff.Equals("WTF")
									&& !daysOff.Equals("TFS")))
									|| (d == 5 && (!daysOff.Equals("TF")
									&& !daysOff.Equals("FS")
									&& !daysOff.Equals("WTF")
									&& !daysOff.Equals("TFS")
									&& !daysOff.Equals("FSS")))
									|| (d == 6 && (!daysOff.Equals("FS")
									&& !daysOff.Equals("SS")
									&& !daysOff.Equals("TFS")
									&& !daysOff.Equals("FSS")
									&& !daysOff.Equals("SSM")))
									)
								{
									vari.SetWorkday(d, workcenter, starttime,
										code, hours);
								} else
                                {
									vari.SetWorkday(d, "", 0, "", 0M);
                                }
							}
							emp.Variations.Add(vari);
							found = true;
						}
						this._dgsc.Employees[e] = emp;
                    }
                }
			}
			return this._dgsc;
		}
	}
}

