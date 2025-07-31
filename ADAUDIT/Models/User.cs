using System;
using System.Collections.Generic; // Added for List

namespace ADAUDIT.Models
{
    public class User
    {
        public string Username { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string DistinguishedName { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Office { get; set; } = string.Empty;
        public string Phone { get; set; } = string.Empty;
        public string Mobile { get; set; } = string.Empty;
        public DateTime? LastLogon { get; set; }
        public DateTime? AccountExpires { get; set; }
        public DateTime? Created { get; set; }
        public DateTime? Modified { get; set; }
        public bool IsEnabled { get; set; }
        public bool PasswordExpired { get; set; }
        public bool AccountLocked { get; set; }
        
        // Additional audit properties
        public bool PasswordNeverExpires { get; set; }
        public bool UserCannotChangePassword { get; set; }
        public bool SmartCardRequired { get; set; }
        public bool TrustedForDelegation { get; set; }
        public bool TrustedToAuthForDelegation { get; set; }
        public bool UseDesKeyOnly { get; set; }
        public bool DontRequirePreauth { get; set; }
        public bool NotDelegated { get; set; }
        public bool RequireSmartCard { get; set; }
        public bool AccountIsSensitive { get; set; }
        public bool DontExpirePassword { get; set; }
        public bool MNSLogonAccount { get; set; }
        public bool TempDuplicateAccount { get; set; }
        public bool NormalAccount { get; set; }
        public bool InterDomainTrustAccount { get; set; }
        public bool WorkstationTrustAccount { get; set; }
        public bool ServerTrustAccount { get; set; }
        public bool PartialSecretsAccount { get; set; }
        
        // Password policy information
        public DateTime? PasswordLastSet { get; set; }
        public DateTime? PasswordExpires { get; set; }
        public int? PasswordAge { get; set; } // in days
        
        // Account lockout information
        public int? BadPasswordCount { get; set; }
        public DateTime? LastBadPasswordAttempt { get; set; }
        public DateTime? LockoutTime { get; set; }
        
        // Logon information
        public int? LogonCount { get; set; }
        public string? LastLogonTimestamp { get; set; }
        public string? PwdLastSet { get; set; }
        
        // Security flags
        public bool IsServiceAccount { get; set; }
        public bool IsAdminAccount { get; set; }
        public bool IsPrivilegedAccount { get; set; }
        public List<string> GroupMemberships { get; set; } = new List<string>();
        public List<string> DirectReports { get; set; } = new List<string>();
        public string Manager { get; set; } = string.Empty;
        
        // Compliance flags
        public bool CompliantWithPasswordPolicy { get; set; }
        public bool CompliantWithAccountPolicy { get; set; }
        public List<string> SecurityIssues { get; set; } = new List<string>();
    }
} 