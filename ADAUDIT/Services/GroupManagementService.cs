using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using ADAUDIT.Models;

namespace ADAUDIT.Services
{
    public class GroupManagementService
    {
        private readonly LDAPService _ldapService;

        public GroupManagementService(LDAPService ldapService)
        {
            _ldapService = ldapService;
        }

        public List<Group> GetAllGroups()
        {
            var groups = new List<Group>();
            
            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(objectClass=group)",
                    new[] { "sAMAccountName", "description", "distinguishedName", "groupType", "member", 
                           "whenCreated", "whenChanged", "memberOf" }
                );

                var results = searcher.FindAll();

                foreach (SearchResult result in results)
                {
                    var group = new Group
                    {
                        Name = GetPropertyValue(result, "sAMAccountName"),
                        Description = GetPropertyValue(result, "description"),
                        DistinguishedName = GetPropertyValue(result, "distinguishedName")
                    };

                    // Parse group type
                    if (result.Properties.Contains("groupType"))
                    {
                        var groupType = Convert.ToInt32(result.Properties["groupType"][0]);
                        group.GroupType = GetGroupTypeDescription(groupType);
                        group.Scope = GetGroupScopeDescription(groupType);
                    }

                    // Parse members
                    if (result.Properties.Contains("member"))
                    {
                        var memberCount = result.Properties["member"].Count;
                        group.MemberCount = memberCount;
                        
                        for (int i = 0; i < memberCount; i++)
                        {
                            var memberDN = result.Properties["member"][i]?.ToString();
                            if (!string.IsNullOrEmpty(memberDN))
                            {
                                group.Members.Add(memberDN);
                            }
                        }
                    }

                    // Parse creation and modification dates
                    if (result.Properties.Contains("whenCreated"))
                    {
                        group.Created = DateTime.Parse(result.Properties["whenCreated"][0].ToString()!);
                    }

                    if (result.Properties.Contains("whenChanged"))
                    {
                        group.Modified = DateTime.Parse(result.Properties["whenChanged"][0].ToString()!);
                    }

                    groups.Add(group);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving groups: {ex.Message}", ex);
            }

            return groups;
        }

        public Group? GetGroupByName(string groupName)
        {
            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    $"(sAMAccountName={groupName})",
                    new[] { "sAMAccountName", "description", "distinguishedName", "groupType", "member", 
                           "whenCreated", "whenChanged", "memberOf" }
                );

                var result = searcher.FindOne();
                if (result == null) return null;

                var group = new Group
                {
                    Name = GetPropertyValue(result, "sAMAccountName"),
                    Description = GetPropertyValue(result, "description"),
                    DistinguishedName = GetPropertyValue(result, "distinguishedName")
                };

                // Parse group type
                if (result.Properties.Contains("groupType"))
                {
                    var groupType = Convert.ToInt32(result.Properties["groupType"][0]);
                    group.GroupType = GetGroupTypeDescription(groupType);
                    group.Scope = GetGroupScopeDescription(groupType);
                }

                // Parse members
                if (result.Properties.Contains("member"))
                {
                    var memberCount = result.Properties["member"].Count;
                    group.MemberCount = memberCount;
                    
                    for (int i = 0; i < memberCount; i++)
                    {
                        var memberDN = result.Properties["member"][i]?.ToString();
                        if (!string.IsNullOrEmpty(memberDN))
                        {
                            group.Members.Add(memberDN);
                        }
                    }
                }

                return group;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving group '{groupName}': {ex.Message}", ex);
            }
        }

        public List<string> GetGroupMembers(string groupName)
        {
            try
            {
                var group = GetGroupByName(groupName);
                return group?.Members ?? new List<string>();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving members for group '{groupName}': {ex.Message}", ex);
            }
        }

        public List<Group> GetGroupsForUser(string username)
        {
            var groups = new List<Group>();
            
            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    $"(&(objectClass=user)(sAMAccountName={username}))",
                    new[] { "memberOf" }
                );

                var result = searcher.FindOne();
                if (result == null) return groups;

                if (result.Properties.Contains("memberOf"))
                {
                    foreach (var groupDN in result.Properties["memberOf"])
                    {
                        var groupName = ExtractGroupNameFromDN(groupDN.ToString()!);
                        var group = GetGroupByName(groupName);
                        if (group != null)
                        {
                            groups.Add(group);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving groups for user '{username}': {ex.Message}", ex);
            }

            return groups;
        }

        private static string GetGroupTypeDescription(int groupType)
        {
            if ((groupType & 0x80000000) != 0) return "Security";
            return "Distribution";
        }

        private static string GetGroupScopeDescription(int groupType)
        {
            var scope = groupType & 0x0000000F;
            return scope switch
            {
                0x00000000 => "Local",
                0x00000001 => "Global",
                0x00000002 => "Domain Local",
                0x00000004 => "Universal",
                _ => "Unknown"
            };
        }

        private static string ExtractGroupNameFromDN(string distinguishedName)
        {
            // Extract the first part of the DN (before the first comma)
            var parts = distinguishedName.Split(',');
            if (parts.Length > 0)
            {
                var firstPart = parts[0];
                if (firstPart.StartsWith("CN="))
                {
                    return firstPart.Substring(3);
                }
            }
            return distinguishedName;
        }

        private static string GetPropertyValue(SearchResult result, string propertyName)
        {
            if (result.Properties.Contains(propertyName) && result.Properties[propertyName].Count > 0)
            {
                return result.Properties[propertyName][0]?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }
    }
} 