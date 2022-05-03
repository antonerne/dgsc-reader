using System.Text.Json;

namespace OsanScheduler.DgscReader.models
{
	public static class DebuggingExtensions
	{
		private static JsonSerializerOptions options
				   = new JsonSerializerOptions { WriteIndented = true };

		public static string ToJson(this object obj)
			=> JsonSerializer.Serialize(obj, options);
	}
}

