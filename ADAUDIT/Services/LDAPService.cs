using System;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.Protocols;
using System.Net;
using System.Security.Principal;

namespace ADAUDIT.Services
{
    public class LDAPService : IDisposable
    {
        private DirectoryEntry? _directoryEntry;
        private PrincipalContext? _principalContext;
        private bool _disposed = false;

        public bool IsConnected => _directoryEntry != null && _principalContext != null;

        public void Connect(string server, string username, string password)
        {
            try
            {
                // Create DirectoryEntry for LDAP operations
                _directoryEntry = new DirectoryEntry($"LDAP://{server}", username, password);

                // Test the connection
                var nativeObject = _directoryEntry.NativeObject;

                // Create PrincipalContext for user management operations
                _principalContext = new PrincipalContext(ContextType.Domain, server, username, password);

                // Test the PrincipalContext connection
                using var testContext = new PrincipalContext(ContextType.Domain, server, username, password);
                if (!testContext.ValidateCredentials(username, password))
                {
                    throw new Exception("Invalid credentials");
                }
            }
            catch (Exception ex)
            {
                Dispose();
                throw new Exception($"Failed to connect to domain controller: {ex.Message}", ex);
            }
        }

        public void Disconnect()
        {
            Dispose();
        }

        public DirectoryEntry GetDirectoryEntry()
        {
            if (_directoryEntry == null)
                throw new InvalidOperationException("Not connected to domain controller");
            return _directoryEntry;
        }

        public PrincipalContext GetPrincipalContext()
        {
            if (_principalContext == null)
                throw new InvalidOperationException("Not connected to domain controller");
            return _principalContext;
        }

        public DirectorySearcher CreateSearcher(string filter = "(objectClass=*)", string[]? propertiesToLoad = null)
        {
            var searcher = new DirectorySearcher(GetDirectoryEntry())
            {
                Filter = filter,
                PageSize = 1000
            };

            if (propertiesToLoad != null)
            {
                searcher.PropertiesToLoad.AddRange(propertiesToLoad);
            }

            return searcher;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                _directoryEntry?.Dispose();
                _principalContext?.Dispose();
                _disposed = true;
            }
        }
    }
} 