using System;
using OsanScheduler.Models.Teams;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;


namespace OsanScheduler.DgscReader.Readers
{
	public class HolidayScheduleReader
	{
		private readonly Team _team;
		private readonly string _file;
		private readonly int _utcDiff;

		public HolidayScheduleReader(Team team, int utc, string file)
		{
			this._file = file;
			this._utcDiff = utc;
			this._team = team;
		}

		public Team Process()
		{
			for (int c = 0; c < this._team.Companies.Count; c++)
            {
				var co = this._team.Companies[c];
				for (int h=0; h < co.Holidays.Count; h++)
                {
					co.Holidays[h].ActualDates.Clear();
                }
				this._team.Companies[c] = co;
            }

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
				var year = (int)row.GetCell(labels["hYear"]).NumericCellValue;
				var code = row.GetCell(labels["Code"]).StringCellValue;
				var company = row.GetCell(labels["Company"]).StringCellValue.ToLower();
				var date = row.GetCell(labels["ActualDate"]).DateCellValue
					.AddHours(this._utcDiff);

				if (year >= previousyear)
                {
					var found = false;
					for (int c=0; c < this._team.Companies.Count && !found; c++)
                    {
						var co = this._team.Companies[c];
						if (co.Code.ToLower().Equals(company))
                        {
							for (int h=0; h < co.Holidays.Count && !found; h++)
                            {
								var hol = co.Holidays[h];
								if (hol.Code.ToLower().Equals(code.ToLower()))
                                {
									found = true;
									hol.AddActualDate(date);
									co.Holidays[h] = hol;
                                }
                            }
							this._team.Companies[c] = co;
                        }
                    }
                }
			}
			return this._team;
		}
	}
}

