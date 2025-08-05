# SonicADRecon

A comprehensive C# Windows Forms application for LDAP domain management, allowing administrators to connect to Active Directory domain controllers, extract users, groups, and group policies, and manage user passwords with appropriate permissions.

## Features

### 🔐 LDAP Connection Management
- Connect to domain controllers using LDAP protocol
- Secure authentication with username/password
- Connection status monitoring
- Support for different domain environments

### 👥 User Management
- **View all domain users** with detailed information:
  - Username and display name
  - Email addresses
  - Last logon times
  - Account status (enabled/disabled, locked, password expired)
  - Department, title, office location
  - Contact information (phone, mobile)
  - Account creation and modification dates

- **Password Management**:
  - Change user passwords with proper permission checks
  - Validation of account status before password changes
  - Support for domain password policies

### 👥 Group Management
- **View all domain groups** with comprehensive details:
  - Group names and descriptions
  - Group types (Security/Distribution)
  - Group scopes (Local/Global/Domain Local/Universal)
  - Member counts and member lists
  - Creation and modification dates

- **Group Analysis**:
  - View group memberships
  - Analyze group hierarchies
  - Track group changes

### 📋 Group Policy Management
- **View all Group Policy Objects (GPOs)**:
  - GPO names and IDs
  - Status (enabled/disabled)
  - Version information
  - Creation and modification dates
  - GPO descriptions and ownership

- **GPO Analysis**:
  - Filter enabled/disabled GPOs
  - View GPOs linked to specific OUs
  - Track GPO versions and changes

## Prerequisites

### System Requirements
- **Operating System**: Windows 10/11 or Windows Server 2016+
- **.NET Framework**: .NET 6.0 or later
- **Domain Access**: Must be a member of the target domain or have network access to the domain controller

### Required Permissions
To use all features of the application, the connecting user should have:

1. **For User Operations**:
   - Read access to user objects
   - Permission to reset user passwords (for password changes)

2. **For Group Operations**:
   - Read access to group objects
   - Permission to view group memberships

3. **For GPO Operations**:
   - Read access to Group Policy objects
   - Permission to view GPO settings

## Installation

### Option 1: Build from Source
1. Clone or download this repository
2. Open the solution in Visual Studio 2022 or later
3. Restore NuGet packages
4. Build the solution (Ctrl+Shift+B)
5. Run the application (F5)

### Option 2: Pre-built Executable
1. Download the latest release from the releases page
2. Extract the ZIP file
3. Run `ADAUDIT.exe`

## Usage

### Initial Setup
1. **Launch the application**
2. **Navigate to the Connection tab**
3. **Enter connection details**:
   - **Domain Controller**: Enter the domain controller name or IP address
     - Example: `dc01.company.local` or `192.168.1.10`
   - **Username**: Enter your domain username
     - Example: `administrator` or `domain\username`
   - **Password**: Enter your domain password
4. **Click "Connect"** to establish the LDAP connection

### Using the Application

#### Connection Tab
- Enter domain controller information
- Test connection status
- View connection feedback

#### Users Tab
- **Refresh Users**: Load all domain users
- **User List**: View comprehensive user information
- **Password Management**: 
  - Select a user from the list
  - Enter new password
  - Click "Change Password"

#### Groups Tab
- **Refresh Groups**: Load all domain groups
- **Group List**: View group details and member counts
- **Group Analysis**: Analyze group structures and memberships

#### Group Policies Tab
- **Refresh GPOs**: Load all Group Policy Objects
- **GPO List**: View GPO details and status
- **GPO Analysis**: Filter and analyze GPO configurations

## Security Considerations

### Authentication
- Credentials are used only for LDAP connection
- Passwords are not stored in the application
- Use domain accounts with minimal required permissions

### Network Security
- LDAP traffic should be encrypted (LDAPS) in production
- Ensure firewall rules allow LDAP traffic (port 389/636)
- Use secure network connections

### Permission Management
- Follow the principle of least privilege
- Use dedicated service accounts for automation
- Regularly audit permissions and access

## Troubleshooting

### Connection Issues
1. **Verify domain controller accessibility**
   - Test network connectivity: `ping dc01.company.local`
   - Check DNS resolution

2. **Verify credentials**
   - Ensure username format is correct
   - Check password validity
   - Verify account is not locked

3. **Check permissions**
   - Ensure user has required AD permissions
   - Verify account is not expired

### Performance Issues
1. **Large domains**: The application loads all objects at once
   - Consider filtering for specific OUs
   - Use pagination for very large environments

2. **Network latency**: Slow connections may cause timeouts
   - Increase timeout settings if needed
   - Use local domain controllers when possible

### Common Error Messages
- **"Invalid credentials"**: Check username/password
- **"Access denied"**: Insufficient permissions
- **"Object not found"**: User/group doesn't exist
- **"Connection timeout"**: Network or firewall issues

## Development

### Project Structure
```
ADAUDIT/
├── ADAUDIT.sln                 # Solution file
├── ADAUDIT/
│   ├── ADAUDIT.csproj         # Project file
│   ├── Program.cs              # Application entry point
│   ├── MainForm.cs            # Main UI form
│   ├── MainForm.Designer.cs   # Form designer
│   ├── Models/                # Data models
│   │   ├── User.cs
│   │   ├── Group.cs
│   │   └── GroupPolicy.cs
│   └── Services/              # Business logic
│       ├── LDAPService.cs
│       ├── UserManagementService.cs
│       ├── GroupManagementService.cs
│       └── GroupPolicyService.cs
└── README.md                  # This file
```

### Dependencies
- **System.DirectoryServices**: Core LDAP functionality
- **System.DirectoryServices.AccountManagement**: User management
- **System.DirectoryServices.Protocols**: Advanced LDAP operations

### Building and Testing
1. Open solution in Visual Studio
2. Restore NuGet packages
3. Build solution
4. Run unit tests (if available)
5. Test with a development domain

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review the error messages
3. Verify your environment setup
4. Create an issue with detailed information

## Version History

- **v1.0.0**: Initial release with basic LDAP functionality
  - User management and password changes
  - Group viewing and analysis
  - GPO listing and status checking
  - Secure LDAP connections 