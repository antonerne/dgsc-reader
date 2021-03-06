using System;
using Microsoft.Extensions.Options;
using MongoDB.Driver;
using MongoDB.Bson;
using OsanScheduler.Models.Employees;
using OsanScheduler.DgscReader.models;

namespace OsanScheduler.DgscReader.Services
{
	public class EmployeeService
	{
		private readonly IMongoCollection<Employee> _empCollection;

		public EmployeeService(ConnectionStrings connectionStrings)
        {
			var mongoClient = new MongoClient(
				connectionStrings.DefaultConnection);
			var mongoDatabase = mongoClient.GetDatabase("scheduler");
			this._empCollection = mongoDatabase.GetCollection<Employee>("employees");
        }

		public async Task<List<Employee>> GetAsync() =>
			await this._empCollection.Find(_ => true).ToListAsync();

		public async Task<List<Employee>> GetByTeamAsync(string id) =>
			await this._empCollection.Find(emp => emp.TeamID.Equals(id)).ToListAsync();

		public async Task<Employee> GetAsync(string id) =>
			await this._empCollection.Find(x => x.Id.Equals(id))
			.FirstOrDefaultAsync();

		public async Task CreateAsync(Employee newEmployee) =>
			await this._empCollection.InsertOneAsync(newEmployee);

		public async Task UpdateAsync(string id, Employee updatedEmployee) =>
			await this._empCollection
			.ReplaceOneAsync(x => x.Id.Equals(id), updatedEmployee);

		public async Task DeleteAsync(string id) =>
			await this._empCollection.DeleteOneAsync(x => x.Id.Equals(id));
	}
}

