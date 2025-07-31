using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using ADAUDIT.Models;

namespace ADAUDIT.Services
{
    public class AuditService
    {
        private readonly LDAPService _ldapService;

        public AuditService(LDAPService ldapService)
        {
            _ldapService = ldapService;
        }

        // New comprehensive audit methods
        public Dictionary<string, object> RunComprehensiveSecurityAudit()
        {
            var auditResults = new Dictionary<string, object>();

            try
            {
                auditResults["Domain Trusts"] = AnalyzeDomainTrusts();
                auditResults["Privileged Groups"] = AnalyzePrivilegedGroupMembership();
                auditResults["Password Policies"] = AnalyzePasswordPolicies();
                auditResults["Kerberos Delegation"] = AnalyzeKerberosDelegation();
                auditResults["AdminSDHolder"] = AnalyzeAdminSDHolderIssues();
                auditResults["Dangerous Rights"] = AnalyzeDangerousRights();
                auditResults["Domain Controllers"] = AnalyzeDomainControllerConfig();
                auditResults["Local Admin Reuse"] = AnalyzeLocalAdminReuse();
                auditResults["DNS Configuration"] = AnalyzeDNSConfiguration();
                auditResults["Stale Objects"] = AnalyzeStaleAndOrphanedObjects();
                auditResults["GPO Risks"] = AnalyzeGPORisks();
                auditResults["LDAP Security"] = AnalyzeLDAPSecurity();
                auditResults["SMB/NTLM Config"] = AnalyzeSMBNTLMConfig();
            }
            catch (Exception ex)
            {
                auditResults["Error"] = $"Comprehensive audit failed: {ex.Message}";
            }

            return auditResults;
        }

        private Dictionary<string, object> AnalyzeDomainTrusts()
        {
            var results = new Dictionary<string, object>();
            var issues = new List<string>();

            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(objectClass=trustedDomain)",
                    new[] { "trustDirection", "trustType", "trustAttributes", "flatName", "name" }
                );

                var trusts = searcher.FindAll();
                results["Total Trusts"] = trusts.Count;

                foreach (SearchResult trust in trusts)
                {
                    var trustName = GetPropertyValue(trust, "name");
                    var trustDirection = GetPropertyValue(trust, "trustDirection");
                    var trustType = GetPropertyValue(trust, "trustType");

                    if (trustDirection == "3") // Bidirectional
                    {
                        issues.Add($"Bidirectional trust with {trustName} - potential security risk");
                    }

                    if (trustType == "2") // External trust
                    {
                        issues.Add($"External trust with {trustName} - review for necessity");
                    }
                }

                results["Issues"] = issues;
                results["Risk Level"] = issues.Count > 0 ? "High" : "Low";
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzePrivilegedGroupMembership()
        {
            var results = new Dictionary<string, object>();
            var privilegedGroups = new List<string> { "Domain Admins", "Enterprise Admins", "Schema Admins", "Account Operators", "Server Operators", "Print Operators", "Backup Operators" };

            try
            {
                foreach (var groupName in privilegedGroups)
                {
                    using var searcher = _ldapService.CreateSearcher(
                        $"(&(objectClass=group)(cn={groupName}))",
                        new[] { "member", "memberOf" }
                    );

                    var group = searcher.FindOne();
                    if (group != null)
                    {
                        var memberCount = group.Properties["member"]?.Count ?? 0;
                        results[groupName] = memberCount;

                        if (memberCount > 10)
                        {
                            results[$"{groupName}_Warning"] = $"High membership count: {memberCount}";
                        }
                    }
                }

                results["Risk Level"] = "Medium";
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzePasswordPolicies()
        {
            var results = new Dictionary<string, object>();

            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(objectClass=domainDNS)",
                    new[] { "minPwdLength", "pwdHistoryLength", "pwdComplexity", "lockoutThreshold", "lockoutDuration", "pwdMaxAge" }
                );

                var domain = searcher.FindOne();
                if (domain != null)
                {
                    var minLength = GetPropertyValue(domain, "minPwdLength");
                    var historyLength = GetPropertyValue(domain, "pwdHistoryLength");
                    var complexity = GetPropertyValue(domain, "pwdComplexity");
                    var lockoutThreshold = GetPropertyValue(domain, "lockoutThreshold");
                    var maxAge = GetPropertyValue(domain, "pwdMaxAge");

                    results["Min Password Length"] = minLength;
                    results["Password History"] = historyLength;
                    results["Complexity Enabled"] = complexity;
                    results["Lockout Threshold"] = lockoutThreshold;
                    results["Max Password Age"] = maxAge;

                    var issues = new List<string>();
                    if (int.TryParse(minLength, out int minLen) && minLen < 12)
                    {
                        issues.Add("Password length below recommended (12 characters)");
                    }

                    if (int.TryParse(historyLength, out int histLen) && histLen < 24)
                    {
                        issues.Add("Password history below recommended (24 passwords)");
                    }

                    if (complexity != "1")
                    {
                        issues.Add("Password complexity not enforced");
                    }

                    results["Issues"] = issues;
                    results["Risk Level"] = issues.Count > 0 ? "Medium" : "Low";
                }
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeKerberosDelegation()
        {
            var results = new Dictionary<string, object>();
            var delegationIssues = new List<string>();

            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=524288))",
                    new[] { "sAMAccountName", "displayName", "userAccountControl", "servicePrincipalName" }
                );

                var delegatedAccounts = searcher.FindAll();
                results["Delegated Accounts Count"] = delegatedAccounts.Count;

                foreach (SearchResult account in delegatedAccounts)
                {
                    var username = GetPropertyValue(account, "sAMAccountName");
                    var displayName = GetPropertyValue(account, "displayName");
                    var spn = GetPropertyValue(account, "servicePrincipalName");

                    if (!string.IsNullOrEmpty(spn))
                    {
                        delegationIssues.Add($"Account {username} ({displayName}) has SPN and delegation enabled");
                    }
                    else
                    {
                        delegationIssues.Add($"Account {username} ({displayName}) has delegation enabled without SPN");
                    }
                }

                results["Issues"] = delegationIssues;
                results["Risk Level"] = delegationIssues.Count > 0 ? "High" : "Low";
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeAdminSDHolderIssues()
        {
            var results = new Dictionary<string, object>();

            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(objectClass=container)",
                    new[] { "cn", "distinguishedName" }
                );

                var containers = searcher.FindAll();
                var adminSDHolder = containers.Cast<SearchResult>()
                    .FirstOrDefault(c => GetPropertyValue(c, "cn") == "AdminSDHolder");

                if (adminSDHolder != null)
                {
                    results["AdminSDHolder Found"] = true;
                    results["Risk Level"] = "Medium";
                    results["Note"] = "AdminSDHolder container exists - review inheritance settings";
                }
                else
                {
                    results["AdminSDHolder Found"] = false;
                    results["Risk Level"] = "Low";
                }
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeDangerousRights()
        {
            var results = new Dictionary<string, object>();
            var dangerousRights = new List<string>();

            try
            {
                // Check for accounts with dangerous rights
                using var searcher = _ldapService.CreateSearcher(
                    "(&(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=524288))",
                    new[] { "sAMAccountName", "displayName" }
                );

                var accountsWithRights = searcher.FindAll();
                results["Accounts with Dangerous Rights"] = accountsWithRights.Count;

                foreach (SearchResult account in accountsWithRights)
                {
                    var username = GetPropertyValue(account, "sAMAccountName");
                    var displayName = GetPropertyValue(account, "displayName");
                    dangerousRights.Add($"{username} ({displayName})");
                }

                results["Dangerous Accounts"] = dangerousRights;
                results["Risk Level"] = dangerousRights.Count > 0 ? "High" : "Low";
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeDomainControllerConfig()
        {
            var results = new Dictionary<string, object>();

            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(&(objectClass=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))",
                    new[] { "dNSHostName", "operatingSystem", "operatingSystemVersion" }
                );

                var domainControllers = searcher.FindAll();
                results["Domain Controllers Count"] = domainControllers.Count;

                var dcInfo = new List<string>();
                foreach (SearchResult dc in domainControllers)
                {
                    var hostname = GetPropertyValue(dc, "dNSHostName");
                    var os = GetPropertyValue(dc, "operatingSystem");
                    var osVersion = GetPropertyValue(dc, "operatingSystemVersion");

                    dcInfo.Add($"{hostname} - {os} {osVersion}");
                }

                results["DC Information"] = dcInfo;
                results["Risk Level"] = "Medium";
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeLocalAdminReuse()
        {
            var results = new Dictionary<string, object>();

            try
            {
                // This would typically require additional tools like BloodHound
                // For now, we'll check for common patterns
                using var searcher = _ldapService.CreateSearcher(
                    "(&(objectClass=user)(sAMAccountName=admin))",
                    new[] { "sAMAccountName", "displayName" }
                );

                var adminAccounts = searcher.FindAll();
                results["Admin Accounts Found"] = adminAccounts.Count;

                var adminList = new List<string>();
                foreach (SearchResult account in adminAccounts)
                {
                    var username = GetPropertyValue(account, "sAMAccountName");
                    var displayName = GetPropertyValue(account, "displayName");
                    adminList.Add($"{username} ({displayName})");
                }

                results["Admin Account List"] = adminList;
                results["Risk Level"] = adminList.Count > 1 ? "High" : "Medium";
                results["Note"] = "Consider implementing LAPS for local admin password management";
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeDNSConfiguration()
        {
            var results = new Dictionary<string, object>();

            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(objectClass=dnsZone)",
                    new[] { "dnsRecord", "zoneType" }
                );

                var zones = searcher.FindAll();
                results["DNS Zones Count"] = zones.Count;

                var zoneInfo = new List<string>();
                foreach (SearchResult zone in zones)
                {
                    var zoneType = GetPropertyValue(zone, "zoneType");
                    zoneInfo.Add($"Zone Type: {zoneType}");
                }

                results["Zone Information"] = zoneInfo;
                results["Risk Level"] = "Medium";
                results["Note"] = "Review DNS zone transfer settings and zone security";
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeStaleAndOrphanedObjects()
        {
            var results = new Dictionary<string, object>();

            try
            {
                // Check for stale user accounts (not logged in for 90+ days)
                var staleUsers = GetInactiveAccounts(90);
                results["Stale User Accounts"] = staleUsers.Count;

                // Check for orphaned groups (no members)
                using var searcher = _ldapService.CreateSearcher(
                    "(&(objectClass=group)(!member=*))",
                    new[] { "cn", "description" }
                );

                var orphanedGroups = searcher.FindAll();
                results["Orphaned Groups"] = orphanedGroups.Count;

                var orphanedGroupList = new List<string>();
                foreach (SearchResult group in orphanedGroups)
                {
                    var groupName = GetPropertyValue(group, "cn");
                    var description = GetPropertyValue(group, "description");
                    orphanedGroupList.Add($"{groupName} - {description}");
                }

                results["Orphaned Group List"] = orphanedGroupList;
                results["Risk Level"] = (staleUsers.Count + orphanedGroups.Count) > 10 ? "Medium" : "Low";
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeGPORisks()
        {
            var results = new Dictionary<string, object>();

            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(objectClass=groupPolicyContainer)",
                    new[] { "displayName", "whenCreated", "whenChanged" }
                );

                var gpos = searcher.FindAll();
                results["Total GPOs"] = gpos.Count;

                var gpoInfo = new List<string>();
                var oldGPOs = 0;

                foreach (SearchResult gpo in gpos)
                {
                    var displayName = GetPropertyValue(gpo, "displayName");
                    var whenCreated = GetPropertyValue(gpo, "whenCreated");
                    var whenChanged = GetPropertyValue(gpo, "whenChanged");

                    if (DateTime.TryParse(whenChanged, out DateTime lastChanged))
                    {
                        var daysSinceChange = (DateTime.Now - lastChanged).TotalDays;
                        if (daysSinceChange > 365)
                        {
                            oldGPOs++;
                            gpoInfo.Add($"{displayName} - Last modified: {daysSinceChange:F0} days ago");
                        }
                    }
                }

                results["Old GPOs (>1 year)"] = oldGPOs;
                results["Old GPO List"] = gpoInfo;
                results["Risk Level"] = oldGPOs > 5 ? "Medium" : "Low";
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeLDAPSecurity()
        {
            var results = new Dictionary<string, object>();

            try
            {
                // Check LDAP signing requirements
                using var searcher = _ldapService.CreateSearcher(
                    "(objectClass=domainDNS)",
                    new[] { "ldapServerIntegrity" }
                );

                var domain = searcher.FindOne();
                if (domain != null)
                {
                    var ldapSigning = GetPropertyValue(domain, "ldapServerIntegrity");
                    results["LDAP Signing"] = ldapSigning;

                    if (ldapSigning != "2") // Not required
                    {
                        results["LDAP Signing Risk"] = "LDAP signing not required - security risk";
                        results["Risk Level"] = "High";
                    }
                    else
                    {
                        results["Risk Level"] = "Low";
                    }
                }
                else
                {
                    results["Risk Level"] = "Unknown";
                }
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        private Dictionary<string, object> AnalyzeSMBNTLMConfig()
        {
            var results = new Dictionary<string, object>();

            try
            {
                // Check for NTLM usage
                using var searcher = _ldapService.CreateSearcher(
                    "(objectClass=domainDNS)",
                    new[] { "msDS-BehaviorVersion" }
                );

                var domain = searcher.FindOne();
                if (domain != null)
                {
                    var behaviorVersion = GetPropertyValue(domain, "msDS-BehaviorVersion");
                    results["Domain Functional Level"] = behaviorVersion;

                    if (int.TryParse(behaviorVersion, out int level) && level < 3)
                    {
                        results["NTLM Risk"] = "Domain functional level below 2008 - NTLM may be required";
                        results["Risk Level"] = "Medium";
                    }
                    else
                    {
                        results["Risk Level"] = "Low";
                    }
                }
                else
                {
                    results["Risk Level"] = "Unknown";
                }
            }
            catch (Exception ex)
            {
                results["Error"] = ex.Message;
            }

            return results;
        }

        public List<User> GetUsersWithSecurityIssues()
        {
            var users = new List<User>();
            
            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(&(objectClass=user)(objectCategory=person))",
                    new[] { 
                        "sAMAccountName", "displayName", "mail", "distinguishedName", "lastLogon", 
                        "accountExpires", "userAccountControl", "department", "title", "physicalDeliveryOfficeName",
                        "telephoneNumber", "mobile", "whenCreated", "whenChanged", "pwdLastSet", "pwdExpires",
                        "lockoutTime", "badPwdCount", "lastLogonTimestamp", "logonCount", "manager",
                        "memberOf", "userPrincipalName", "servicePrincipalName"
                    }
                );

                var results = searcher.FindAll();

                foreach (SearchResult result in results)
                {
                    var user = ParseUserWithAuditData(result);
                    AnalyzeUserSecurity(user);
                    users.Add(user);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving users for audit: {ex.Message}", ex);
            }

            return users;
        }

        private User ParseUserWithAuditData(SearchResult result)
        {
            var user = new User
            {
                Username = GetPropertyValue(result, "sAMAccountName"),
                DisplayName = GetPropertyValue(result, "displayName"),
                Email = GetPropertyValue(result, "mail"),
                DistinguishedName = GetPropertyValue(result, "distinguishedName"),
                Department = GetPropertyValue(result, "department"),
                Title = GetPropertyValue(result, "title"),
                Office = GetPropertyValue(result, "physicalDeliveryOfficeName"),
                Phone = GetPropertyValue(result, "telephoneNumber"),
                Mobile = GetPropertyValue(result, "mobile"),
                Manager = GetPropertyValue(result, "manager")
            };

            // Parse user account control flags
            if (result.Properties.Contains("userAccountControl"))
            {
                var uac = Convert.ToInt32(result.Properties["userAccountControl"][0]);
                ParseUserAccountControl(user, uac);
            }

            // Parse password information
            ParsePasswordInformation(user, result);

            // Parse lockout information
            ParseLockoutInformation(user, result);

            // Parse group memberships
            ParseGroupMemberships(user, result);

            // Parse dates
            ParseDateInformation(user, result);

            return user;
        }

        private void ParseUserAccountControl(User user, int uac)
        {
            user.IsEnabled = (uac & 0x2) == 0; // ACCOUNTDISABLE
            user.PasswordExpired = (uac & 0x800000) != 0; // PASSWORD_EXPIRED
            user.AccountLocked = (uac & 0x10) != 0; // LOCKOUT
            user.PasswordNeverExpires = (uac & 0x10000) != 0; // DONT_EXPIRE_PASSWORD
            user.UserCannotChangePassword = (uac & 0x40) != 0; // PASSWD_CANT_CHANGE
            user.SmartCardRequired = (uac & 0x40000) != 0; // SMARTCARD_REQUIRED
            user.TrustedForDelegation = (uac & 0x80000) != 0; // TRUSTED_FOR_DELEGATION
            user.TrustedToAuthForDelegation = (uac & 0x1000000) != 0; // TRUSTED_TO_AUTH_FOR_DELEGATION
            user.UseDesKeyOnly = (uac & 0x200000) != 0; // USE_DES_KEY_ONLY
            user.DontRequirePreauth = (uac & 0x400000) != 0; // DONT_REQUIRE_PREAUTH
            user.NotDelegated = (uac & 0x100000) != 0; // NOT_DELEGATED
            user.RequireSmartCard = (uac & 0x40000) != 0; // SMARTCARD_REQUIRED
            user.AccountIsSensitive = (uac & 0x1000000) != 0; // TRUSTED_TO_AUTH_FOR_DELEGATION
            user.DontExpirePassword = (uac & 0x10000) != 0; // DONT_EXPIRE_PASSWORD
            user.MNSLogonAccount = (uac & 0x20000) != 0; // MNS_LOGON_ACCOUNT
            user.TempDuplicateAccount = (uac & 0x800) != 0; // TEMP_DUPLICATE_ACCOUNT
            user.NormalAccount = (uac & 0x200) != 0; // NORMAL_ACCOUNT
            user.InterDomainTrustAccount = (uac & 0x800) != 0; // INTERDOMAIN_TRUST_ACCOUNT
            user.WorkstationTrustAccount = (uac & 0x1000) != 0; // WORKSTATION_TRUST_ACCOUNT
            user.ServerTrustAccount = (uac & 0x2000) != 0; // SERVER_TRUST_ACCOUNT
            user.PartialSecretsAccount = (uac & 0x4000000) != 0; // PARTIAL_SECRETS_ACCOUNT
        }

        private void ParsePasswordInformation(User user, SearchResult result)
        {
            if (result.Properties.Contains("pwdLastSet"))
            {
                var pwdLastSetValue = result.Properties["pwdLastSet"][0];
                if (pwdLastSetValue != null)
                {
                    try
                    {
                        if (pwdLastSetValue is long ticks && ticks > 0)
                        {
                            user.PasswordLastSet = DateTime.FromFileTime(ticks);
                            user.PasswordAge = (int)(DateTime.Now - user.PasswordLastSet.Value).TotalDays;
                        }
                    }
                    catch (Exception)
                    {
                        user.PasswordLastSet = null;
                    }
                }
            }

            if (result.Properties.Contains("pwdExpires"))
            {
                var pwdExpiresValue = result.Properties["pwdExpires"][0];
                if (pwdExpiresValue != null)
                {
                    try
                    {
                        if (pwdExpiresValue is long ticks && ticks > 0 && ticks != 0x7FFFFFFFFFFFFFFF)
                        {
                            user.PasswordExpires = DateTime.FromFileTime(ticks);
                        }
                    }
                    catch (Exception)
                    {
                        user.PasswordExpires = null;
                    }
                }
            }
        }

        private void ParseLockoutInformation(User user, SearchResult result)
        {
            if (result.Properties.Contains("lockoutTime"))
            {
                var lockoutTimeValue = result.Properties["lockoutTime"][0];
                if (lockoutTimeValue != null)
                {
                    try
                    {
                        if (lockoutTimeValue is long ticks && ticks > 0)
                        {
                            user.LockoutTime = DateTime.FromFileTime(ticks);
                        }
                    }
                    catch (Exception)
                    {
                        user.LockoutTime = null;
                    }
                }
            }

            if (result.Properties.Contains("badPwdCount"))
            {
                var badPwdCountValue = result.Properties["badPwdCount"][0];
                if (badPwdCountValue != null)
                {
                    try
                    {
                        user.BadPasswordCount = Convert.ToInt32(badPwdCountValue);
                    }
                    catch (Exception)
                    {
                        user.BadPasswordCount = null;
                    }
                }
            }
        }

        private void ParseGroupMemberships(User user, SearchResult result)
        {
            if (result.Properties.Contains("memberOf"))
            {
                foreach (var group in result.Properties["memberOf"])
                {
                    if (group != null)
                    {
                        var groupDn = group.ToString()!;
                        var groupName = groupDn.Split(',')[0].Replace("CN=", "");
                        user.GroupMemberships.Add(groupName);
                        
                        // Check for privileged groups
                        if (groupName.Contains("Administrators") || groupName.Contains("Domain Admins") || 
                            groupName.Contains("Enterprise Admins") || groupName.Contains("Schema Admins"))
                        {
                            user.IsAdminAccount = true;
                            user.IsPrivilegedAccount = true;
                        }
                        
                        if (groupName.Contains("Service") || groupName.Contains("SPN"))
                        {
                            user.IsServiceAccount = true;
                        }
                    }
                }
            }
        }

        private void ParseDateInformation(User user, SearchResult result)
        {
            // Parse last logon
            if (result.Properties.Contains("lastLogon"))
            {
                var lastLogonValue = result.Properties["lastLogon"][0];
                if (lastLogonValue != null)
                {
                    try
                    {
                        if (lastLogonValue is long ticks && ticks > 0)
                        {
                            user.LastLogon = DateTime.FromFileTime(ticks);
                        }
                    }
                    catch (Exception)
                    {
                        user.LastLogon = null;
                    }
                }
            }

            // Parse account expiration
            if (result.Properties.Contains("accountExpires"))
            {
                var expiresValue = result.Properties["accountExpires"][0];
                if (expiresValue != null)
                {
                    try
                    {
                        if (expiresValue is long ticks && ticks > 0 && ticks != 0x7FFFFFFFFFFFFFFF)
                        {
                            user.AccountExpires = DateTime.FromFileTime(ticks);
                        }
                    }
                    catch (Exception)
                    {
                        user.AccountExpires = null;
                    }
                }
            }

            // Parse creation and modification dates
            if (result.Properties.Contains("whenCreated"))
            {
                try
                {
                    var createdValue = result.Properties["whenCreated"][0];
                    if (createdValue != null)
                    {
                        user.Created = DateTime.Parse(createdValue.ToString()!);
                    }
                }
                catch (Exception)
                {
                    user.Created = null;
                }
            }

            if (result.Properties.Contains("whenChanged"))
            {
                try
                {
                    var changedValue = result.Properties["whenChanged"][0];
                    if (changedValue != null)
                    {
                        user.Modified = DateTime.Parse(changedValue.ToString()!);
                    }
                }
                catch (Exception)
                {
                    user.Modified = null;
                }
            }
        }

        private void AnalyzeUserSecurity(User user)
        {
            var issues = new List<string>();

            // Password policy compliance
            if (user.PasswordNeverExpires)
            {
                issues.Add("Password never expires");
                user.CompliantWithPasswordPolicy = false;
            }

            if (user.PasswordAge.HasValue && user.PasswordAge > 90)
            {
                issues.Add($"Password age: {user.PasswordAge} days (should be < 90)");
                user.CompliantWithPasswordPolicy = false;
            }

            if (user.PasswordExpired)
            {
                issues.Add("Password is expired");
                user.CompliantWithPasswordPolicy = false;
            }

            // Account security
            if (!user.IsEnabled)
            {
                issues.Add("Account is disabled");
            }

            if (user.AccountLocked)
            {
                issues.Add("Account is locked");
            }

            if (user.BadPasswordCount.HasValue && user.BadPasswordCount > 5)
            {
                issues.Add($"Multiple failed logon attempts: {user.BadPasswordCount}");
            }

            // Privileged account checks
            if (user.IsAdminAccount)
            {
                issues.Add("User has administrative privileges");
            }

            if (user.IsServiceAccount)
            {
                issues.Add("Service account detected");
            }

            // Delegation security
            if (user.TrustedForDelegation)
            {
                issues.Add("Account trusted for delegation");
            }

            if (user.TrustedToAuthForDelegation)
            {
                issues.Add("Account trusted to authenticate for delegation");
            }

            // Smart card requirements
            if (user.SmartCardRequired && !user.RequireSmartCard)
            {
                issues.Add("Smart card required but not enforced");
            }

            // Last logon analysis
            if (user.LastLogon.HasValue)
            {
                var daysSinceLastLogon = (DateTime.Now - user.LastLogon.Value).TotalDays;
                if (daysSinceLastLogon > 90)
                {
                    issues.Add($"Inactive account: {daysSinceLastLogon:F0} days since last logon");
                }
            }
            else
            {
                issues.Add("Never logged on");
            }

            user.SecurityIssues = issues;
            user.CompliantWithAccountPolicy = issues.Count == 0;
        }

        public List<User> GetUsersWithWeakPasswords()
        {
            return GetUsersWithSecurityIssues()
                .Where(u => !u.CompliantWithPasswordPolicy)
                .ToList();
        }

        public List<User> GetPrivilegedAccounts()
        {
            return GetUsersWithSecurityIssues()
                .Where(u => u.IsPrivilegedAccount || u.IsAdminAccount)
                .ToList();
        }

        public List<User> GetInactiveAccounts(int daysThreshold = 90)
        {
            return GetUsersWithSecurityIssues()
                .Where(u => u.LastLogon.HasValue && 
                           (DateTime.Now - u.LastLogon.Value).TotalDays > daysThreshold)
                .ToList();
        }

        public List<User> GetLockedAccounts()
        {
            return GetUsersWithSecurityIssues()
                .Where(u => u.AccountLocked)
                .ToList();
        }

        public List<User> GetServiceAccounts()
        {
            return GetUsersWithSecurityIssues()
                .Where(u => u.IsServiceAccount)
                .ToList();
        }

        public Dictionary<string, int> GetSecurityIssueSummary()
        {
            var users = GetUsersWithSecurityIssues();
            var summary = new Dictionary<string, int>();

            summary["Total Users"] = users.Count;
            summary["Weak Passwords"] = users.Count(u => !u.CompliantWithPasswordPolicy);
            summary["Privileged Accounts"] = users.Count(u => u.IsPrivilegedAccount);
            summary["Locked Accounts"] = users.Count(u => u.AccountLocked);
            summary["Disabled Accounts"] = users.Count(u => !u.IsEnabled);
            summary["Service Accounts"] = users.Count(u => u.IsServiceAccount);
            summary["Inactive Accounts (>90 days)"] = users.Count(u => 
                u.LastLogon.HasValue && (DateTime.Now - u.LastLogon.Value).TotalDays > 90);

            return summary;
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