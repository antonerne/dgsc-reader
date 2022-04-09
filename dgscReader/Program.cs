using System.Collections.Generic;
using System.IO;
using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Teams;
using OsanScheduler.Models.Employees.Labor;
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

// get the location for the data files
Console.WriteLine("Scheduler Files Loc ({0}):", settings.DefaultDataDirectory);
var loc = Console.ReadLine();
if (loc != null && !loc.Equals(""))
{
    settings.DefaultDataDirectory = loc;
}



// Create service classes for data
TeamsService teamService = new TeamsService(settings);
EmployeeService empService = new EmployeeService(settings);
WorksService worksService = new WorksService(settings);

// Read the team(s), finding the DFS Team, then find the DGS-C Site.
Team dfs = await teamService.GetByCodeAsync("dfs");
Site dgsc = dfs.Sites.First<Site>(s => s.Code == "dgsc");
var previousYear = DateTime.UtcNow.Year - 1;

// Now get the list of available Employees for the team.
var emps = new List<ObjectId>();
if (dfs.Id != ObjectId.Empty)
{
    var employees = await empService.GetByTeamAsync(dfs.Id);
    employees.ForEach(emp =>
    {
        var current = emp.GetCurrentSite(new DateTime(previousYear, 1, 1), DateTime.UtcNow);
        if (dgsc.ID.Equals(current) || dgsc.Code.ToLower().Equals(current.ToLower()))
        {
            dgsc.Employees.Add(emp);
            emps.Add(emp.Id);
        }
    });
}
var work = await worksService.GetAsync();
work.ForEach(wk =>
{
    var found = false;

    for (int e=0; e < dgsc.Employees.Count && !found; e++)
    {
        var emp = dgsc.Employees[e];
        if (emp.Id.Equals(wk.EmployeeId))
        {
            emp.Work.Add(wk);
        }
        dgsc.Employees[e] = emp;
    }
});

// get the list of excel files in the designated directory
var files = Directory.GetFiles(settings.DefaultDataDirectory, "*.xlsx");

// find the employee file and process using EmployeeReader class object
string empFile = "";
string siteLaborFile = "";
foreach (var file in files)
{
    if (file.ToLower().EndsWith("employees.xlsx")
        && !file.ToLower().EndsWith("~$employees.xlsx"))
    {
        empFile = file;
    } else if (file.ToLower().EndsWith("laborcodes.xlsx")
        && !file.ToLower().EndsWith("~$laborcodes.xlsx"))
    {
        siteLaborFile = file;
    }
}

if (!empFile.Equals(""))
{
    Console.WriteLine(empFile);
    EmployeesReader reader = new EmployeesReader(dgsc, dfs.Id, empFile);
    dgsc = reader.Process();
}
if (!siteLaborFile.Equals(""))
{
    Console.WriteLine(siteLaborFile);
    SiteLaborCodeReader reader = new SiteLaborCodeReader(dgsc, siteLaborFile);
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
            case "employeelaborcodes.xlsx":
                Console.WriteLine(file);
                var eReader = new EmpLaborCodeReader(dgsc, file);
                dgsc = eReader.Process();
                break;
            case "holidayschedule.xlsx":
                Console.WriteLine(file);
                var hReader = new HolidayScheduleReader(dfs, dgsc.UtcDifference,
                    file);
                dfs = hReader.Process();
                break;
            case "leaves.xlsx":
                Console.WriteLine(file);
                var lReader = new LeavesReader(dgsc, file);
                dgsc = lReader.Process();
                break;
            case "schedulevariations.xlsx":
                Console.WriteLine(file);
                var vReader = new VariationReader(dgsc, file);
                dgsc = vReader.Process();
                break;
            case "workhours.xlsx":
                Console.WriteLine(file);
                var wReader = new WorkReader(dgsc, file);
                dgsc = wReader.Process();
                break;
            case "workschedule.xlsx":
                Console.WriteLine(file);
                var sReader = new ScheduleReader(dgsc, file);
                dgsc = sReader.Process();
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
        List<Work> newWorks = new List<Work>();
        emp.Work.ForEach(async (wk) =>
        {
            if (wk.Id != ObjectId.Empty)
            {
                await worksService.UpdateAsync(wk.Id, wk);
            }
            else
            {
                newWorks.Add(wk);
            }
        });
        if (newWorks.Count > 0)
        {
            await worksService.CreateManyAsync(newWorks);
        }
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

// Update the dgsc site in the team and the teams object to the database
var found = false;
for (int s=0; s < dfs.Sites.Count && !found; s++)
{
    if (dfs.Sites[s].Equals(dgsc))
    {
        dfs.Sites[s] = dgsc;
        found = true;
    }
}
await teamService.UpdateAsync(dfs.Id, dfs);

Console.WriteLine("Completed");

await host.RunAsync();
