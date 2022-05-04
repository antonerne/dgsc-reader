using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using OsanScheduler.Models.DB;
using OsanScheduler.Models.Sites;
using OsanScheduler.Models.Teams;
using OsanScheduler.Models.Employees.Labor;
using OsanScheduler.DgscReader.models;
using OsanScheduler.DgscReader.Readers;
using Microsoft.EntityFrameworkCore;
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

using var context = new SchedulerDBContext(settings.DefaultConnection);

// get the location for the data files
Console.WriteLine("Scheduler Files Loc ({0}):", settings.DefaultDataDirectory);
var loc = Console.ReadLine();
if (loc != null && !loc.Equals(""))
{
    settings.DefaultDataDirectory = loc;
}

if (context.dbTeams.Count() <= 0)
{
    // load initial data
    string initialloc = "/Users/antonerne/Projects/scheduler/csharp/dgsc-reader/dgscReader/initial.json";
    using (StreamReader r = new StreamReader(initialloc))
    {
        string json = r.ReadToEnd();
        List<Team>? cts = JsonSerializer.Deserialize<List<Team>>(json);
        if (cts != null)
        {
            Console.WriteLine(cts.Count);
            cts.ForEach(tm =>
            {
                Console.WriteLine(tm.ToJson());
                tm.DateCreated = DateTime.UtcNow;
                tm.LastUpdated = DateTime.UtcNow;
                if (tm.Sites != null)
                {
                    tm.Sites.ForEach(site =>
                    {
                        site.LastUpdated = DateTime.UtcNow;
                        site.DateCreated = DateTime.UtcNow;
                    });
                }
                context.dbTeams.Add(tm);
                Console.WriteLine("Added {0}", tm.Name);
            });
            context.SaveChanges();
        }
    }
}

// Read the team(s), finding the DFS Team, then find the DGS-C Site.
Team dfs = context.dbTeams
    .Include(tm => tm.Companies)
    .ThenInclude(co => co.Holidays)
    .Include(tm => tm.ContactTypes)
    .Include(tm => tm.DisplayCodes)
    .Include(tm => tm.SpecialtyGroups)
    .ThenInclude(sg => sg.Areas)
    .Include(tm => tm.Sites)
    .ThenInclude(site => site.LaborCodes)
    .Include(tm => tm.Sites)
    .ThenInclude(site => site.WorkCodes)
    .Include(tm => tm.Sites)
    .ThenInclude(site => site.Workcenters)
    .ThenInclude(wc => wc.Positions)
    .Single(tm => tm.Code.ToLower().Equals("dfs"));
    
Console.WriteLine(dfs.ToJson());
Site dgsc = dfs.Sites.First<Site>(s => s.Code == "dgsc");
var previousYear = DateTime.UtcNow.Year - 1;

// Now get the list of available Employees for the team.
var emps = new List<int>();
if (dfs.Id != 0)
{
    var employees = context.dbEmployees
        .Where(e => e.TeamID == dfs.Id).ToList();
    employees.ForEach(emp =>
    {
        var current = emp.GetCurrentSite(new DateTime(previousYear, 1, 1), DateTime.UtcNow);
        if (dgsc.Code.ToLower().Equals(current.ToLower()))
        {
            dgsc.Employees.Add(emp);
            emps.Add(emp.Id);
        }
    });
}
var empWork = context.dbWork.ToList();
empWork.ForEach(ewk =>
{
    var found = false;

    for (int e=0; e < dgsc.Employees.Count && !found; e++)
    {
        var emp = dgsc.Employees[e];
        if (emp.Id.Equals(ewk.Id))
        {
            emp.Work.Add(ewk);
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
dgsc.Employees.ForEach(emp =>
{
    if (emp.Id > 0)
    {
        context.dbEmployees.Update(emp);
    } else
    {
        context.dbEmployees.Add(emp);
    }
    context.SaveChanges();
    Console.WriteLine(emp.ToJson());

    /*if (emp.Work.Count > 0)
    {
        if (emp.Id.Equals(""))
        {
            context.dbEmployees.Add(emp);
            context.SaveChanges();
        }
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
    }*/
});

    // Update the dgsc site in the team and the teams object to the database
    /*var found = false;
    for (int s=0; s < dfs.Sites.Count && !found; s++)
    {
    if (dfs.Sites[s].Equals(dgsc))
    {
        dfs.Sites[s] = dgsc;
        found = true;
    }
    
    await teamService.UpdateAsync(dfs.Id, dfs);
    */

Console.WriteLine("Completed");

host.Run();
