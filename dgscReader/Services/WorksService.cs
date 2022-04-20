using System;
using Microsoft.Extensions.Options;
using MongoDB.Driver;
using MongoDB.Bson;
using OsanScheduler.Models.Employees.Labor;
using OsanScheduler.DgscReader.models;

namespace OsanScheduler.DgscReader.Services
{
	public class WorksService
	{
		private readonly IMongoCollection<EmployeeWork> _workCollection;

		public WorksService(ConnectionStrings connectionStrings)
		{
			var mongoClient = new MongoClient(
				connectionStrings.DefaultConnection);
			var mongoDatabase = mongoClient.GetDatabase("scheduler");
			this._workCollection = mongoDatabase.GetCollection<EmployeeWork>("works");
		}

		public async Task<List<EmployeeWork>> GetAsync() =>
			await this._workCollection.Find(_ => true).ToListAsync();

		public async Task<EmployeeWork> GetAsync(string id) =>
			await this._workCollection.Find(x => x.Id == id)
			.FirstOrDefaultAsync();

		public async Task CreateAsync(EmployeeWork newWork) =>
			await this._workCollection.InsertOneAsync(newWork);

		public async Task CreateManyAsync(List<EmployeeWork> newWorks) =>
			await this._workCollection.InsertManyAsync(newWorks);

		public async Task UpdateAsync(string id, EmployeeWork updated) =>
			await this._workCollection
			.ReplaceOneAsync(x => x.Id == id, updated);

		public async Task DeleteAsync(string id) =>
			await this._workCollection.DeleteOneAsync(x => x.Id == id);
	}
}

