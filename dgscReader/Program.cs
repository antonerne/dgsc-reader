using System.Collections.Generic;
using System.IO;
using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Teams;
using OsanScheduler.Models.Employees;
using MongoDB.Driver;
using MongoDB.Bson;
using OsanScheduler.DgscReader.models;
using OsanScheduler.DgscReader.Services;
using OsanScheduler.DgscReader.Readers;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

// get default settings
using IHost host = Host.CreateDefaultBuilder(args).Build();

IConfiguration config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .AddEnvironmentVariables()
    .Build();

ConnectionStrings settings = config.GetRequiredSection("ConnectionStrings")
    .Get<ConnectionStrings>();


// Create service classes for data
TeamsService teamService = new TeamsService(settings);
EmployeeService empService = new EmployeeService(settings);
WorksService worksService = new WorksService(settings);

// Read the team(s), finding the DFS Team, then find the DGS-C Site.
Team dfs = await teamService.GetByCodeAsync("dfs");
Site dgsc = dfs.Sites.First<Site>(s => s.Code == "dgsc");
var previousYear = DateTime.UtcNow.Year - 1;

// Now get the list of available Employees for the team.
if (dfs.Id != ObjectId.Empty)
{
    var employees = await empService.GetByTeamAsync(dfs.Id);
    employees.ForEach(emp =>
    {
        var current = emp.GetCurrentSite(new DateTime(previousYear, 1, 1), DateTime.UtcNow);
        if (dgsc.ID.Equals(current) || dgsc.Code.ToLower().Equals(current.ToLower()))
        {
            dgsc.Employees.Add(emp);
        }
    });
}

// get the list of excel files in the designated directory
var files = Directory.GetFiles(settings.DefaultDataDirectory, "*.xlsx");

// find the employee file and process using EmployeeReader class object
string empFile = "";
foreach (var file in files)
{
    if (file.ToLower().EndsWith("employees.xlsx")
        && !file.ToLower().EndsWith("~$employees.xlsx"))
    {
        empFile = file;
    }
}

if (!empFile.Equals(""))
{
    Console.WriteLine(empFile);
    EmployeesReader reader = new EmployeesReader(dgsc, dfs.Id, empFile);
    dgsc = reader.Process();
}

// process the rest of the excel files and process them to fill the employee
// information
foreach (var file in files)
{
    var filename = file.Split("/").Last();
    if (filename != null && !filename.StartsWith("~"))
    {
        switch (filename.ToLower())
        {
            case "annualleave.xlsx":
                Console.WriteLine(file);
                var aReader = new AnnualLeaveReader(dgsc, file);
                dgsc = aReader.Process();
                break;
        }
    }
}


// lastly update the team, site employees, and employee work information to the
// database.
dgsc.Employees.ForEach(async (emp) =>
{
    if (emp.Work.Count > 0)
    {
        emp.Work.ForEach(async (wk) =>
        {
            if (wk.Id != null && !wk.Id.Equals(""))
            {
                await worksService.UpdateAsync(wk.Id, wk);
            }
            else
            {
                await worksService.CreateAsync(wk);
            }
        });
    }
    try
    {
        if (emp.Id != ObjectId.Empty)
        {
            await empService.UpdateAsync(emp.Id, emp);
        }
        else
        {
            await empService.CreateAsync(emp);
        }
    } catch (Exception ex)
    {
        Console.WriteLine(ex.StackTrace);
    }
});

Console.WriteLine("Completed");

await host.RunAsync();
