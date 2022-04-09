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
		private readonly IMongoCollection<Work> _workCollection;

		public WorksService(ConnectionStrings connectionStrings)
		{
			var mongoClient = new MongoClient(
				connectionStrings.DefaultConnection);
			var mongoDatabase = mongoClient.GetDatabase("scheduler");
			this._workCollection = mongoDatabase.GetCollection<Work>("works");
		}

		public async Task<List<Work>> GetAsync() =>
			await this._workCollection.Find(_ => true).ToListAsync();

		public async Task<Work> GetAsync(string id) =>
			await this._workCollection.Find(x => x.Id == new ObjectId(id))
			.FirstOrDefaultAsync();

		public async Task CreateAsync(Work newWork) =>
			await this._workCollection.InsertOneAsync(newWork);

		public async Task UpdateAsync(ObjectId id, Work updated) =>
			await this._workCollection
			.ReplaceOneAsync(x => x.Id == id, updated);

		public async Task DeleteAsync(ObjectId id) =>
			await this._workCollection.DeleteOneAsync(x => x.Id == id);
	}
}

