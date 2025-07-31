using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using ADAUDIT.Models;

namespace ADAUDIT.Services
{
    public class UserManagementService
    {
        private readonly LDAPService _ldapService;

        public UserManagementService(LDAPService ldapService)
        {
            _ldapService = ldapService;
        }

        public List<User> GetAllUsers()
        {
            var users = new List<User>();
            
            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    "(&(objectClass=user)(objectCategory=person))",
                    new[] { "sAMAccountName", "displayName", "mail", "distinguishedName", "lastLogon", 
                           "accountExpires", "userAccountControl", "department", "title", "physicalDeliveryOfficeName",
                           "telephoneNumber", "mobile", "whenCreated", "whenChanged" }
                );

                var results = searcher.FindAll();

                foreach (SearchResult result in results)
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
                        Mobile = GetPropertyValue(result, "mobile")
                    };

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
                                else if (lastLogonValue is string strValue && !string.IsNullOrEmpty(strValue))
                                {
                                    // Try parsing as string if it's not a long
                                    if (long.TryParse(strValue, out long parsedTicks) && parsedTicks > 0)
                                    {
                                        user.LastLogon = DateTime.FromFileTime(parsedTicks);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                // Ignore parsing errors for lastLogon
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
                                else if (expiresValue is string strValue && !string.IsNullOrEmpty(strValue))
                                {
                                    if (long.TryParse(strValue, out long parsedTicks) && parsedTicks > 0 && parsedTicks != 0x7FFFFFFFFFFFFFFF)
                                    {
                                        user.AccountExpires = DateTime.FromFileTime(parsedTicks);
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                // Ignore parsing errors for accountExpires
                                user.AccountExpires = null;
                            }
                        }
                    }

                    // Parse user account control flags
                    if (result.Properties.Contains("userAccountControl"))
                    {
                        var uac = Convert.ToInt32(result.Properties["userAccountControl"][0]);
                        user.IsEnabled = (uac & 0x2) == 0; // ACCOUNTDISABLE flag
                        user.PasswordExpired = (uac & 0x800000) != 0; // PASSWORD_EXPIRED flag
                        user.AccountLocked = (uac & 0x10) != 0; // LOCKOUT flag
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
                            // Ignore parsing errors for whenCreated
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
                            // Ignore parsing errors for whenChanged
                            user.Modified = null;
                        }
                    }

                    users.Add(user);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving users: {ex.Message}", ex);
            }

            return users;
        }

        public void ChangePassword(string username, string newPassword)
        {
            try
            {
                using var context = _ldapService.GetPrincipalContext();
                using var user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
                
                if (user == null)
                {
                    throw new Exception($"User '{username}' not found");
                }

                // Check if user account is enabled
                if (!user.Enabled.GetValueOrDefault())
                {
                    throw new Exception($"User '{username}' account is disabled");
                }

                // Check if user account is locked
                if (user.IsAccountLockedOut())
                {
                    throw new Exception($"User '{username}' account is locked");
                }

                // Set the new password
                user.SetPassword(newPassword);
                user.Save();

                // Force password change on next logon (optional)
                // user.ExpirePasswordNow();
                // user.Save();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error changing password for user '{username}': {ex.Message}", ex);
            }
        }

        public User? GetUserByUsername(string username)
        {
            try
            {
                using var searcher = _ldapService.CreateSearcher(
                    $"(&(objectClass=user)(sAMAccountName={username}))",
                    new[] { "sAMAccountName", "displayName", "mail", "distinguishedName", "lastLogon", 
                           "accountExpires", "userAccountControl", "department", "title", "physicalDeliveryOfficeName",
                           "telephoneNumber", "mobile", "whenCreated", "whenChanged" }
                );

                var result = searcher.FindOne();
                if (result == null) return null;

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
                    Mobile = GetPropertyValue(result, "mobile")
                };

                // Parse additional properties similar to GetAllUsers
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
                            else if (lastLogonValue is string strValue && !string.IsNullOrEmpty(strValue))
                            {
                                if (long.TryParse(strValue, out long parsedTicks) && parsedTicks > 0)
                                {
                                    user.LastLogon = DateTime.FromFileTime(parsedTicks);
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // Ignore parsing errors for lastLogon
                            user.LastLogon = null;
                        }
                    }
                }

                if (result.Properties.Contains("userAccountControl"))
                {
                    var uac = Convert.ToInt32(result.Properties["userAccountControl"][0]);
                    user.IsEnabled = (uac & 0x2) == 0;
                    user.PasswordExpired = (uac & 0x800000) != 0;
                    user.AccountLocked = (uac & 0x10) != 0;
                }

                return user;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving user '{username}': {ex.Message}", ex);
            }
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