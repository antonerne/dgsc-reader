using System;
using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Employees.Assignments;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace OsanScheduler.DgscReader.Readers
{
	public class ScheduleReader
	{
		private readonly Site _dgsc;
		private readonly string _file;

		public ScheduleReader(Site dgsc, string file)
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
				var schID = (int)row.GetCell(labels["SortID"]).NumericCellValue;
				var schedule = row.GetCell(labels["Schedule"]).StringCellValue;
				var days = schedule.Split("|");
				var found = false;

				// find the employee, then get the last assignment in the list.
				// get the employee's workcenter, if the assignment is lacking
				// the number of schedules, add until enough are present.
				for (int e=0; e < this._dgsc.Employees.Count && !found; e++)
                {
					var emp = this._dgsc.Employees[e];
					if (emp.CompanyInfo.CompanyID.Equals(empID))
                    {
						emp.Assignments.Sort();
						var asgmt = emp.Assignments[emp.Assignments.Count - 1];
						while (schID > asgmt.Schedules.Count - 1)
                        {
							asgmt.AddSchedule(7);
                        }
						asgmt.Schedules.Sort();
						var wkctr = "";
						for (int w=0; w < asgmt.Schedules[0].Workdays.Count
							&& wkctr.Equals(""); w++)
                        {
							if (!asgmt.Schedules[0].Workdays[w].Workcenter
								.Equals(""))
                            {
								wkctr = asgmt.Schedules[0].Workdays[w].Workcenter;
                            }
                        }
						// workhour will be either 8 or 10 depending on the
						// number of workdays with a code.
						decimal hours = 8M;
						int count = 0;
						for (int d=0; d < days.Length; d++)
                        {
							if (!days[d].Equals("")) count++;
                        }
						if (count < 5) hours = 10M;
						var sch = asgmt.Schedules[schID];
						for (int d=0; d < days.Length; d++)
                        {
							var day = d - 1;
							if (day < 0) day = 6;
							int starttime = -1;
							if (!days[d].Equals(""))
                            {
								this._dgsc.WorkCodes.ForEach(wc =>
								{
									if (wc.Code.Equals(days[d]))
                                    {
										starttime = wc.StartTime;
                                    }
								});
								sch.SetWorkday(day, wkctr, starttime, days[d],
									hours);
                            } else
                            {
								sch.SetWorkday(day, "", 0, "", 0);
                            }
                        }
						asgmt.Schedules[schID] = sch;
						emp.Assignments[emp.Assignments.Count - 1] = asgmt;
						this._dgsc.Employees[e] = emp;
						found = true;
                    }
                }
			}
			return this._dgsc;
		}
	}
}
