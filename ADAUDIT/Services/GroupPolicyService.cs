using System;
using System.Collections.Generic;
using System.DirectoryServices;
using ADAUDIT.Models;

namespace ADAUDIT.Services
{
    public class GroupPolicyService
    {
        private readonly LDAPService _ldapService;

        public GroupPolicyService(LDAPService ldapService)
        {
            _ldapService = ldapService;
        }

        public List<GroupPolicy> GetAllGPOs()
        {
            var gpos = new List<GroupPolicy>();
            
            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(objectClass=groupPolicyContainer)",
                    new[] { "displayName", "name", "distinguishedName", "whenCreated", "whenChanged", 
                           "gPCFileSysPath", "flags", "versionNumber" }
                );

                var results = searcher.FindAll();

                foreach (SearchResult result in results)
                {
                    var gpo = new GroupPolicy
                    {
                        Name = GetPropertyValue(result, "displayName"),
                        Id = GetPropertyValue(result, "name"),
                        DistinguishedName = GetPropertyValue(result, "distinguishedName"),
                        Description = GetPropertyValue(result, "description")
                    };

                    // Parse GPO status based on flags
                    if (result.Properties.Contains("flags"))
                    {
                        var flags = Convert.ToInt32(result.Properties["flags"][0]);
                        gpo.IsEnabled = (flags & 0x1) == 0; // GPO_DISABLE flag
                        gpo.Status = gpo.IsEnabled ? "Enabled" : "Disabled";
                    }
                    else
                    {
                        gpo.IsEnabled = true;
                        gpo.Status = "Enabled";
                    }

                    // Parse version
                    if (result.Properties.Contains("versionNumber"))
                    {
                        var version = Convert.ToInt32(result.Properties["versionNumber"][0]);
                        gpo.Version = version.ToString();
                    }

                    // Parse creation and modification dates
                    if (result.Properties.Contains("whenCreated"))
                    {
                        gpo.Created = DateTime.Parse(result.Properties["whenCreated"][0].ToString()!);
                    }

                    if (result.Properties.Contains("whenChanged"))
                    {
                        gpo.Modified = DateTime.Parse(result.Properties["whenChanged"][0].ToString()!);
                    }

                    gpos.Add(gpo);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving GPOs: {ex.Message}", ex);
            }

            return gpos;
        }

        public GroupPolicy? GetGPOById(string gpoId)
        {
            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    $"(name={gpoId})",
                    new[] { "displayName", "name", "distinguishedName", "whenCreated", "whenChanged", 
                           "gPCFileSysPath", "flags", "versionNumber" }
                );

                var result = searcher.FindOne();
                if (result == null) return null;

                var gpo = new GroupPolicy
                {
                    Name = GetPropertyValue(result, "displayName"),
                    Id = GetPropertyValue(result, "name"),
                    DistinguishedName = GetPropertyValue(result, "distinguishedName"),
                    Description = GetPropertyValue(result, "description")
                };

                // Parse GPO status based on flags
                if (result.Properties.Contains("flags"))
                {
                    var flags = Convert.ToInt32(result.Properties["flags"][0]);
                    gpo.IsEnabled = (flags & 0x1) == 0; // GPO_DISABLE flag
                    gpo.Status = gpo.IsEnabled ? "Enabled" : "Disabled";
                }
                else
                {
                    gpo.IsEnabled = true;
                    gpo.Status = "Enabled";
                }

                // Parse version
                if (result.Properties.Contains("versionNumber"))
                {
                    var version = Convert.ToInt32(result.Properties["versionNumber"][0]);
                    gpo.Version = version.ToString();
                }

                return gpo;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving GPO '{gpoId}': {ex.Message}", ex);
            }
        }

        public List<GroupPolicy> GetGPOsForOU(string ouDistinguishedName)
        {
            var gpos = new List<GroupPolicy>();
            
            try
            {
                // Search for GPOs linked to the specified OU
                using var searcher = _ldapService.CreateSearcher(
                    $"(objectClass=organizationalUnit)",
                    new[] { "gPLink" }
                );

                var result = searcher.FindOne();
                if (result != null && result.Properties.Contains("gPLink"))
                {
                    var gplink = result.Properties["gPLink"][0]?.ToString();
                    if (!string.IsNullOrEmpty(gplink))
                    {
                        // Parse the gPLink attribute which contains linked GPOs
                        var gpoLinks = ParseGPLink(gplink);
                        foreach (var gpoId in gpoLinks)
                        {
                            var gpo = GetGPOById(gpoId);
                            if (gpo != null)
                            {
                                gpos.Add(gpo);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving GPOs for OU '{ouDistinguishedName}': {ex.Message}", ex);
            }

            return gpos;
        }

        public List<GroupPolicy> GetEnabledGPOs()
        {
            var allGPOs = GetAllGPOs();
            return allGPOs.FindAll(gpo => gpo.IsEnabled);
        }

        public List<GroupPolicy> GetDisabledGPOs()
        {
            var allGPOs = GetAllGPOs();
            return allGPOs.FindAll(gpo => !gpo.IsEnabled);
        }

        private static List<string> ParseGPLink(string gplink)
        {
            var gpoIds = new List<string>();
            
            // gPLink format: [LDAP://cn={GUID},cn=policies,cn=system,DC=domain,DC=com;0][LDAP://cn={GUID},cn=policies,cn=system,DC=domain,DC=com;1]
            var links = gplink.Split(']', StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var link in links)
            {
                if (link.StartsWith("[LDAP://cn={"))
                {
                    var startIndex = link.IndexOf("{") + 1;
                    var endIndex = link.IndexOf("}");
                    if (startIndex > 0 && endIndex > startIndex)
                    {
                        var gpoId = link.Substring(startIndex, endIndex - startIndex);
                        gpoIds.Add(gpoId);
                    }
                }
            }
            
            return gpoIds;
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