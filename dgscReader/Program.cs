using System.Collections.Generic;
using System.IO;
using System.Text.Json;
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


// Create service classes for data
TeamsService teamService = new TeamsService(settings);
EmployeeService empService = new EmployeeService(settings);
WorksService worksService = new WorksService(settings);

// get the location for the data files
Console.WriteLine("Scheduler Files Loc ({0}):", settings.DefaultDataDirectory);
var loc = Console.ReadLine();
if (loc != null && !loc.Equals(""))
{
    settings.DefaultDataDirectory = loc;
}

// add initial data if no teams are present
List<Team> teams = await teamService.GetAsync();
if (teams == null || teams.Count <= 0)
{
    // no teams, so no initial data loaded.  Add it from the initial data
    // json file
    using (StreamReader r = new StreamReader(settings.InitialDataLocation))
    {
        string json = r.ReadToEnd();
        List<Team>? cts = JsonSerializer.Deserialize<List<Team>>(json);
        if (cts != null)
        {
            cts.ForEach(async ct =>
            {
                ct.Id = Guid.NewGuid().ToString();
                ct.DateCreated = DateTime.UtcNow;
                ct.LastUpdated = DateTime.UtcNow;
                if (ct.Sites != null && ct.Sites.Count > 0)
                {
                    ct.Sites.ForEach(site =>
                    {
                        site.ID = Guid.NewGuid().ToString();
                        site.DateCreated = DateTime.UtcNow;
                        site.LastUpdated = DateTime.UtcNow;
                    });
                }
                await teamService.CreateAsync(ct);
            });
        }
    }
}


// Read the team(s), finding the DFS Team, then find the DGS-C Site.
Team dfs = await teamService.GetByCodeAsync("dfs");
Site dgsc = dfs.Sites.First<Site>(s => s.Code == "dgsc");
var previousYear = DateTime.UtcNow.Year - 1;

// Now get the list of available Employees for the team.
var emps = new List<string>();
if (dfs.Id != "")
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
var empWork = await worksService.GetAsync();
empWork.ForEach(ewk =>
{
    var found = false;

    for (int e=0; e < dgsc.Employees.Count && !found; e++)
    {
        var emp = dgsc.Employees[e];
        if (emp.Id.Equals(ewk.Id))
        {
            ewk.Work.ForEach(wk =>
            {
                emp.Work.Add(wk);
            });
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
        if (emp.Id.Equals(""))
        {
            emp.Id = Guid.NewGuid().ToString();
        }
        EmployeeWork eWork = new EmployeeWork();
        eWork.Id = emp.Id;
        emp.Work.ForEach(wk =>
        {
            if (wk.Id.Equals(""))
            {
                wk.Id = Guid.NewGuid().ToString();
            }
            eWork.Work.Add(wk);
        });
        var ew = await worksService.GetAsync(eWork.Id);
        if (ew != null)
        {
            await worksService.UpdateAsync(eWork.Id, eWork);
        } else
        {
            await worksService.CreateAsync(eWork);
        }
    }
    try
    {
        var e = await empService.GetAsync(emp.Id);
        if (e != null)
        {
            await empService.UpdateAsync(e.Id, emp);
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
