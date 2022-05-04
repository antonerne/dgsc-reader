using System;
namespace OsanScheduler.DgscReader.models
{
	public class ConnectionStrings
	{
		public string DefaultConnection { get; set; } = null!;
		public string DefaultDataDirectory { get; set; } = null!;
		public string InitialDataLocation { get; set; } = default!;
	}
}

