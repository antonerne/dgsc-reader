using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson.Serialization.IdGenerators;
using OsanScheduler.Models.Employees;
using OsanScheduler.Models.Teams.CompanyInfo;
using OsanScheduler.Models.Teams.TeamInfo;
using OsanScheduler.Models.Teams;
using OsanScheduler.Models.Sites;

namespace OsanScheduler.DgscReader.models
{
    public class Team2
    {
#nullable enable
        [BsonId]
        [BsonRepresentation(BsonType.String)]
        [JsonPropertyName("id")]
        public string Id { get; set; }
#nullable disable
        [BsonElement("code")]
        [JsonPropertyName("code")]
        public string Code { get; set; }
        [BsonElement("name")]
        [JsonPropertyName("name")]
        public string Name { get; set; }
        [BsonElement("date_created")]
        [JsonPropertyName("date_created")]
        public DateTime DateCreated { get; set; }
        [BsonElement("last_updated")]
        [JsonPropertyName("last_updated")]
        public DateTime LastUpdated { get; set; }
        [BsonElement("companies")]
        [JsonPropertyName("companies")]
        public List<Company> Companies { get; set; }
        [BsonElement("displayCodes"), BsonIgnoreIfNull]
        [JsonPropertyName("displayCodes")]
        public List<DisplayCode> DisplayCodes { get; set; }
        [BsonElement("specialtyGroups"), BsonIgnoreIfNull]
        [JsonPropertyName("specialtyGroups")]
        public List<SpecialtyGroup> SpecialtyGroups { get; set; }
        [BsonElement("contactTypes"), BsonIgnoreIfNull]
        [JsonPropertyName("contactTypes")]
        public List<Contact> ContactTypes { get; set; }
        [BsonElement("permissions"), BsonIgnoreIfNull]
        [JsonPropertyName("permissions")]
        public List<Permission> Permissions { get; set; }
        [BsonElement("sites"), BsonIgnoreIfNull]
        [JsonPropertyName("sites")]
        public List<Site> Sites { get; set; }

        public Team2()
        {
            this.Id = Guid.NewGuid().ToString();
            this.Code = "";
            this.Name = "";
            this.DateCreated = new DateTime(0);
            this.LastUpdated = new DateTime(0);
            this.Companies = new List<Company>();
            this.DisplayCodes = new List<DisplayCode>();
            this.SpecialtyGroups = new List<SpecialtyGroup>();
            this.ContactTypes = new List<Contact>();
            this.Permissions = new List<Permission>();
            this.Sites = new List<Site>();
        }
        public Team2(Team team)
        {
            this.Id = team.Id.ToString();
            this.Code = team.Code;
            this.Name = team.Name;
            this.DateCreated = new DateTime(team.DateCreated.Ticks);
            this.LastUpdated = new DateTime(team.LastUpdated.Ticks);
            this.Companies = new List<Company>();
            team.Companies.ForEach(co =>
            {
                this.Companies.Add(new Company(co));
            });
            this.DisplayCodes = new List<DisplayCode>();
            team.DisplayCodes.ForEach(dc =>
            {
                this.DisplayCodes.Add(new DisplayCode(dc));
            });
            this.SpecialtyGroups = new List<SpecialtyGroup>();
            team.SpecialtyGroups.ForEach(sg =>
            {
                this.SpecialtyGroups.Add(new SpecialtyGroup(sg));
            });
            this.ContactTypes = new List<Contact>();
            team.ContactTypes.ForEach(ct =>
            {
                this.ContactTypes.Add(new Contact(ct));
            });
            this.Permissions = new List<Permission>();
            team.Permissions.ForEach(perm =>
            {
                this.Permissions.Add(new Permission(perm));
            });
            this.Sites = new List<Site>();
            team.Sites.ForEach(site =>
            {
                this.Sites.Add(new Site(site));
            });
        }

        public Company AddCompany(string code, string title, string timecard)
        {
            Company answer = null;
            this.Companies.ForEach(co =>
            {
                if (co.Equals(code, title))
                {
                    answer = co;
                    co.TimeCardSystem = timecard;
                }
            });
            if (answer == null)
            {
                answer = new Company
                {
                    Code = code,
                    Title = title,
                    TimeCardSystem = timecard
                };
                this.Companies.Add(answer);
            }
            this.LastUpdated = DateTime.UtcNow;
            return answer;
        }

        public void RemoveCompany(string id, string code)
        {
            if (!id.Equals(""))
            {
                this.Companies.RemoveAll(co => co.ID.Equals(id));
            } else
            {
                this.Companies.RemoveAll(co => co.Equals(code));
            }
            this.LastUpdated = DateTime.UtcNow;
        }

        public DisplayCode AddDisplayCode(string code, string name,
            string bkcolor, string txcolor, bool isLv)
        {
            var order = -1;
            DisplayCode answer = null;
            this.DisplayCodes.ForEach(dc =>
            {
                if (dc.Equals(code))
                {
                    answer = dc;
                    dc.Name = name;
                    dc.BackColor = bkcolor;
                    dc.TextColor = txcolor;
                    dc.IsLeave = isLv;
                }
                else if (dc.Order > order)
                {
                    order = dc.Order;
                }
            });
            if (answer == null)
            {
                answer = new DisplayCode
                {
                    Code = code,
                    Name = name,
                    BackColor = bkcolor,
                    TextColor = txcolor,
                    IsLeave = isLv,
                    Order = order + 1
                };
                this.DisplayCodes.Add(answer);
            }
            this.LastUpdated = DateTime.UtcNow;
            return answer;
        }

        public void RemoveDisplayCode(string id, string code)
        {
            if (!id.Equals(""))
            {
                this.DisplayCodes.RemoveAll(dc => dc.ID.Equals(id));
            } else
            {
                this.DisplayCodes.RemoveAll(dc => dc.Equals(code));
            }
            this.DisplayCodes.Sort();
            for (int i=0; i < this.DisplayCodes.Count; i++)
            {
                this.DisplayCodes[i].Order = i;
            }
            this.LastUpdated = DateTime.UtcNow;
        }

        public void SwapDisplayCode(string id, string direction)
        {
            this.DisplayCodes.Sort();
            var found = -1;
            for (int i = 0; i < this.DisplayCodes.Count && found < 0; i++)
            {
                if (this.DisplayCodes[i].ID.Equals(id))
                {
                    found = i;
                }
            }
            if (found >= 0)
            {
                if (direction.ToLower().Substring(0, 1).Equals("u") && found > 0)
                {
                    var temp = this.DisplayCodes[found].Order;
                    this.DisplayCodes[found].Order = this.DisplayCodes[found - 1].Order;
                    this.DisplayCodes[found - 1].Order = temp;
                }
                else if (direction.ToLower().Substring(0, 1).Equals("d")
                  && found < this.DisplayCodes.Count - 1)
                {
                    var temp = this.DisplayCodes[found].Order;
                    this.DisplayCodes[found].Order = this.DisplayCodes[found + 1].Order;
                    this.DisplayCodes[found + 1].Order = temp;
                }
            }
        }

        public SpecialtyGroup AddSpecialtyGroup(string code, string title)
        {
            SpecialtyGroup answer = null;
            this.SpecialtyGroups.ForEach(sg =>
            {
                if (sg.Equals(code))
                {
                    answer = sg;
                    sg.Title = title;
                    this.LastUpdated = DateTime.UtcNow;
                }
            });
            if (answer == null)
            {
                answer = new SpecialtyGroup
                {
                    Code = code,
                    Title = title
                };
                this.SpecialtyGroups.Add(answer);
                this.LastUpdated = DateTime.UtcNow;
            }
            return answer;
        }

        public void RemoveSpecialtyGroup(string id, string code)
        {
            if (!id.Equals(""))
            {
                this.SpecialtyGroups.RemoveAll(sg => sg.ID.Equals(id));
            } else
            {
                this.SpecialtyGroups.RemoveAll(sg => sg.Equals(code));
            }
            this.LastUpdated = DateTime.UtcNow;
        }

        public Contact AddContactType(string code, string desc, bool req)
        {
            var order = -1;
            Contact answer = null;
            this.ContactTypes.ForEach(ct =>
            {
                if (ct.Equals(code))
                {
                    ct.Description = desc;
                    ct.IsRequired = req;
                    answer = ct;
                } else if (ct.Order > order)
                {
                    order = ct.Order;
                }
            });
            if (answer == null)
            {
                answer = new Contact
                {
                    Code = code,
                    Description = desc,
                    IsRequired = req,
                    Order = order + 1
                };
                this.ContactTypes.Add(answer);
            }
            this.LastUpdated = DateTime.UtcNow;
            return answer;
        }

        public void RemoveContactType(string id, string code)
        {
            if (!id.Equals(""))
            {
                this.ContactTypes.RemoveAll(ct => ct.ID.Equals(id));
            } else
            {
                this.ContactTypes.RemoveAll(ct => ct.Equals(code));
            }
            for (int i=0; i < this.ContactTypes.Count; i++)
            {
                this.ContactTypes[i].Order = i;
            }
        }

        public void SwapContactType(string id, string direction)
        {
            this.ContactTypes.Sort();
            var found = -1;
            for (int i = 0; i < this.ContactTypes.Count && found < 0; i++)
            {
                if (this.ContactTypes[i].ID.Equals(id))
                {
                    found = i;
                }
            }
            if (found >= 0)
            {
                if (direction.ToLower().Substring(0, 1).Equals("u") && found > 0)
                {
                    var temp = this.ContactTypes[found].Order;
                    this.ContactTypes[found].Order = this.ContactTypes[found - 1].Order;
                    this.ContactTypes[found - 1].Order = temp;
                }
                else if (direction.ToLower().Substring(0, 1).Equals("d")
                  && found < this.DisplayCodes.Count - 1)
                {
                    var temp = this.ContactTypes[found].Order;
                    this.ContactTypes[found].Order = this.ContactTypes[found + 1].Order;
                    this.ContactTypes[found + 1].Order = temp;
                }
            }
            this.LastUpdated = DateTime.UtcNow;
        }

        public Permission AddPermission(string title, PermissionLevel lvl,
            bool read, bool write, bool approve, bool admin)
        {
            Permission answer = null;
            this.Permissions.ForEach(p =>
            {
                if (p.Title.ToLower().Equals(title.ToLower()))
                {
                    answer = p;
                    p.Level = lvl;
                    p.Read = read;
                    p.Write = write;
                    p.Approver = approve;
                    p.Admin = admin;
                }
            });
            if (answer == null)
            {
                answer = new Permission
                {
                    Title = title,
                    Level = lvl,
                    Read = read,
                    Write = write,
                    Approver = approve,
                    Admin = admin
                };
                this.Permissions.Add(answer);
            }
            this.LastUpdated = DateTime.UtcNow;
            return answer;
        }

        public void RemovePermission(string id, string title)
        {
            if (!id.Equals(""))
            {
                this.Permissions.RemoveAll(p => p.ID.Equals(id));
            } else
            {
                this.Permissions.RemoveAll(p => p.Equals(title));
            }
            this.LastUpdated = DateTime.UtcNow;
        }

        public Site AddSite(string code, string title, int utcdiff)
        {
            Site answer = null;
            this.Sites.ForEach(site =>
            {
                if (site.Equals(code))
                {
                    answer = site;
                    site.Title = title;
                    site.UtcDifference = utcdiff;
                }
            });
            if (answer == null)
            {
                answer = new Site
                {
                    Code = code,
                    Title = title,
                    UtcDifference = utcdiff
                };
                this.Sites.Add(answer);
            }
            this.LastUpdated = DateTime.UtcNow;
            return answer;
        }

        public void RemoveSite(string id, string code)
        {
            if (!id.Equals(""))
            {
                this.Sites.RemoveAll(s => s.ID.Equals(id));
            } else
            {
                this.Sites.RemoveAll(s => s.Equals(code));
            }
            this.LastUpdated = DateTime.UtcNow;
        }
    }
}
