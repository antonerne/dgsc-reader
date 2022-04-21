using System;
using Microsoft.Extensions.Options;
using MongoDB.Driver;
using MongoDB.Bson;
using OsanScheduler.Models.Teams;
using OsanScheduler.DgscReader.models;

namespace OsanScheduler.DgscReader.Services
{
	public class TeamsService
	{
		private readonly IMongoCollection<Team> _teamsCollection;
		private readonly IMongoCollection<Team2> _teams2Collection;

		public TeamsService(ConnectionStrings connectionStrings)
        {
			var mongoClient = new MongoClient(
				connectionStrings.DefaultConnection);
			var mongoDatabase = mongoClient.GetDatabase("scheduler");
			this._teamsCollection = mongoDatabase.GetCollection<Team>("teams");
			this._teams2Collection = mongoDatabase.GetCollection<Team2>("teams");
        }

		public async Task<List<Team>> GetAsync() =>
			await this._teamsCollection.Find(_ => true).ToListAsync();

		public async Task<Team> GetAsync(string id) =>
			await this._teamsCollection.Find(x => x.Id == new ObjectId(id))
			.FirstOrDefaultAsync();

		public async Task<Team2> GetByCodeAsync(string code) =>
			await this._teams2Collection.Find(x => x.Code == code)
			.FirstOrDefaultAsync();

		public async Task CreateAsync(Team newTeam) =>
			await this._teamsCollection.InsertOneAsync(newTeam);

		public async Task CreateManyAsync(Team2[] newTeams) =>
			await this._teams2Collection.InsertManyAsync(newTeams);

		public async Task UpdateAsync(ObjectId id, Team updatedTeam) =>
			await this._teamsCollection
			.ReplaceOneAsync(x => x.Id == id, updatedTeam);

		public async Task UpdateAsync(string id, Team2 updatedTeam) =>
			await this._teams2Collection
			.ReplaceOneAsync(x => x.Id == id, updatedTeam);

		public async Task DeleteAsync(ObjectId id) =>
			await this._teamsCollection.DeleteOneAsync(x => x.Id == id);
	}
}

