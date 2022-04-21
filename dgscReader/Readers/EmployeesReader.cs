using System;
using System.IO;
using System.Collections.Generic;
using MongoDB.Bson;
using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Employees;
using OsanScheduler.Models.Employees.Assignments;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace OsanScheduler.DgscReader.Readers
{
	public class EmployeesReader
	{
		private readonly Site _dgsc;
		private readonly string _file;
		private readonly string _teamid;

		public EmployeesReader(Site dgsc, string teamid, string empFile)
		{
			this._dgsc = dgsc;
			this._teamid = teamid;
			this._file = empFile;
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

			for (int i=0; i < row.Cells.Count; i++)
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

			for (int i=1; i <= sheet.LastRowNum; i++)
            {
				row = sheet.GetRow(i);
				var endDate = row.GetCell(labels["EndDate"]).DateCellValue;
				var startDate = row.GetCell(labels["StartDate"]).DateCellValue;
				while (startDate.DayOfWeek != DayOfWeek.Sunday)
                {
					startDate = startDate.AddDays(-1);
                }
				if (endDate.Equals(baseDate) || endDate >= expireDate)
				{
					var empID = row.GetCell(labels["EmployeeID"]).StringCellValue;
					var lastName = row.GetCell(labels["LastName"]).StringCellValue;
					var company = row.GetCell(labels["Company"]).StringCellValue;
					var firstName = row.GetCell(labels["FirstName"]).StringCellValue;
					var midCell = row.GetCell(labels["MiddleName"]);
					var middleName = "";
					if (midCell != null && midCell.CellType != CellType.Blank)
					{
						middleName = midCell.StringCellValue;
					}
					var jobTitle = "";
					midCell = row.GetCell(labels["JobTitle"]);
					if (midCell != null && midCell.CellType != CellType.Blank)
					{
						jobTitle = midCell.StringCellValue;
					}
					var workcenter = row.GetCell(labels["WorkCenter"]).StringCellValue;
					int schedChangeFreq = (int)row.GetCell(labels["ScheduleChangeFreq"])
						.NumericCellValue;
					var schedChangeDate = row.GetCell(labels["ScheduleChangeDate"])
						.DateCellValue;
					while (schedChangeDate.DayOfWeek != DayOfWeek.Sunday)
                    {
						schedChangeDate = schedChangeDate.AddDays(1);
                    }
					var altID = "";
					midCell = row.GetCell(labels["PeoplesoftID"]);
					if (midCell != null && midCell.CellType != CellType.Blank)
					{
						altID = midCell.StringCellValue;
					}
					var rank = "";
					midCell = row.GetCell(labels["LaborCategory"]);
					if (midCell != null && midCell != null && midCell.CellType != CellType.Blank)
					{
						rank = midCell.StringCellValue;
					}
					var division = "";
					midCell = row.GetCell(labels["SubCompany"]);
					if (midCell != null && midCell.CellType != CellType.Blank)
					{
						division = midCell.StringCellValue;
					}
					var costcenter = "";
					midCell = row.GetCell(labels["CostCenter"]);
					if (midCell != null && midCell.CellType != CellType.Blank)
					{
						costcenter = midCell.StringCellValue;
					}

					// now that I've pulled the data for the employee, I need to
					// use it to either create a new employee or to update the
					// employee's record at the site.
					Employee? emp = this._dgsc.Employees
						.Find(e => e.CompanyInfo.EmployeeID.Equals(empID));
					if (emp != null)
					{
						// employee listed for site, so update.
						int empPos = -1;
						for (int e = 0; e < this._dgsc.Employees.Count
							&& empPos < 0; e++)
                        {
							if (this._dgsc.Employees[e].CompanyInfo.EmployeeID.Equals(empID))
                            {
								empPos = e;
                            }
						}
						emp.Name.Last = lastName;
						emp.Name.First = firstName;
						emp.Name.Middle = middleName;
						emp.CompanyInfo.CompanyCode = company.ToLower();
						emp.CompanyInfo.EmployeeID = empID;
						emp.CompanyInfo.AlternateID = altID;
						emp.CompanyInfo.Rank = rank;
						emp.CompanyInfo.Division = division;
						emp.CompanyInfo.CostCenter = costcenter;

						// set the assignments to match the data.

						// reset enddate, if it is the basedate.
						if (endDate.Equals(baseDate))
                        {
							endDate = new DateTime(9999, 12, 31);
                        }

						// if scheduleChangeFreq > 0, then ensure there are
						// two assignments, else only one.
						emp.Assignments.Sort();
						if (schedChangeFreq > 0)
                        {
							while (emp.Assignments.Count > 2)
                            {
								emp.Assignments.RemoveAt(2);
                            }
							var asgmt = emp.Assignments[0];
							asgmt.StartDate = startDate;
							asgmt.EndDate = schedChangeDate.AddDays(-1);
							asgmt.Site = this._dgsc.Code;
							asgmt.JobTitle = jobTitle;
							asgmt.DaysInRotation = 0;
							asgmt.Schedules.Sort();
							while (asgmt.Schedules.Count > 1)
                            {
								asgmt.Schedules.RemoveAt(1);
                            }
							if (asgmt.Schedules.Count == 0)
                            {
								asgmt.AddSchedule(7);
                            }
							for (int d = 1; d < 6; d++)
							{
								asgmt.Schedules[0].SetWorkday(d, workcenter, 6, "D", 8M);
							}
							asgmt = emp.Assignments[1];
							asgmt.StartDate = schedChangeDate;
							asgmt.EndDate = endDate;
							asgmt.Site = this._dgsc.Code;
							asgmt.JobTitle = jobTitle;
							asgmt.DaysInRotation = schedChangeFreq;
							asgmt.Schedules.Sort();
							while (asgmt.Schedules.Count > 1)
							{
								asgmt.Schedules.RemoveAt(1);
							}
							if (asgmt.Schedules.Count == 0)
							{
								asgmt.AddSchedule(7);
							}
							for (int d = 1; d < 6; d++)
							{
								asgmt.Schedules[0].SetWorkday(d, workcenter, 6, "D", 8M);
							}
						} else
                        {
							while (emp.Assignments.Count > 1)
                            {
								emp.Assignments.RemoveAt(1);
                            }
							var asgmt = emp.Assignments[0];
							asgmt.StartDate = startDate;
							asgmt.EndDate = endDate;
							asgmt.Site = this._dgsc.Code;
							asgmt.JobTitle = jobTitle;
							asgmt.DaysInRotation = 0;
							asgmt.Schedules.Sort();
							while (asgmt.Schedules.Count > 1)
							{
								asgmt.Schedules.RemoveAt(1);
							}
							if (asgmt.Schedules.Count == 0)
							{
								asgmt.AddSchedule(7);
							}
							for (int d = 1; d < 6; d++)
							{
								asgmt.Schedules[0].SetWorkday(d, workcenter, 6, "D", 8M);
							}
						}
						this._dgsc.Employees[empPos] = emp;
					}
					else
					{
						// employee not list for site, so add.
						emp = new Employee();
						emp.TeamID = this._teamid;
						emp.Name.Last = lastName;
						emp.Name.First = firstName;
						emp.Name.Middle = middleName;
						emp.Email = firstName.ToLower() + "."
							+ lastName.ToLower() + "@" + company.ToLower()
							+ ".com";
						emp.CompanyInfo.CompanyCode = company.ToLower();
						emp.CompanyInfo.EmployeeID = empID;
						emp.CompanyInfo.AlternateID = altID;
						emp.CompanyInfo.Rank = rank;
						emp.CompanyInfo.Division = division;
						emp.CompanyInfo.CostCenter = costcenter;
						emp.Roles.Add("Employee");
						emp.Creds.SetPassword("InitialPassword");
						emp.Creds.MustChange = true;

						// add assignments for start and end dates, but if
						// schedulechangefreq > 0 then need to add second
						// assignment starting on the schedulechangedate.
						// if enddate is 1/1/1970 (baseDate) change to 12/31/9999.
						if (endDate.Equals(baseDate))
						{
							endDate = new DateTime(9999, 12, 31);
						}
						if (schedChangeFreq > 0)
						{
							// add an assignment for the period of start to the
							// day before the schedulechangedate, then add a
							// an assignment for the period of schedulechangedate
							// to enddate, unless the schedulechange date is
							// after enddate
							if (schedChangeDate > endDate)
							{
								var asgmt = new Assignment();
								asgmt.StartDate = startDate;
								asgmt.EndDate = endDate;
								asgmt.Site = this._dgsc.Code;
								asgmt.JobTitle = jobTitle;
								asgmt.DaysInRotation = 0;
								asgmt.AddSchedule(7);
								for (int d = 1; d < 6; d++)
								{
									asgmt.Schedules[0].SetWorkday(d, workcenter,
										6, "D", 8M);
								}
								emp.Assignments.Add(asgmt);
							}
							else
							{
								var asgmt = new Assignment();
								asgmt.StartDate = startDate;
								asgmt.EndDate = schedChangeDate.AddDays(-1);
								asgmt.Site = this._dgsc.Code;
								asgmt.JobTitle = jobTitle;
								asgmt.DaysInRotation = 0;
								asgmt.AddSchedule(7);
								for (int d = 1; d < 6; d++)
								{
									asgmt.Schedules[0].SetWorkday(d, workcenter, 6, "D", 8M);
								}
								emp.Assignments.Add(asgmt);
								asgmt = new Assignment();
								asgmt.StartDate = schedChangeDate;
								asgmt.EndDate = endDate;
								asgmt.Site = this._dgsc.Code;
								asgmt.JobTitle = jobTitle;
								asgmt.DaysInRotation = schedChangeFreq;
								asgmt.AddSchedule(7);
								for (int d = 1; d < 6; d++)
								{
									asgmt.Schedules[0].SetWorkday(d, workcenter, 6, "D", 8M);
								}
								emp.Assignments.Add(asgmt);
							}
						}
						else
						{
							// single assignment with start and enddates, a
							// single schedule for the assignment with workdays
							// temporarily scheduled for M-F, Day Shift (6-2)
							var asgmt = new Assignment();
							asgmt.StartDate = startDate;
							asgmt.EndDate = endDate;
							asgmt.Site = this._dgsc.Code;
							asgmt.JobTitle = jobTitle;
							asgmt.DaysInRotation = 0;
							asgmt.AddSchedule(7);
							for (int d = 1; d < 6; d++)
							{
								asgmt.Schedules[0].SetWorkday(d, workcenter, 6,
									"D", 8M);
							}
							emp.Assignments.Add(asgmt);
						}
						this._dgsc.Employees.Add(emp);
					}
				}
            }

			return this._dgsc;
        }
	}
}

