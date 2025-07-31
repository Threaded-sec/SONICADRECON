# Quick Setup Guide - SonicADRecon

## 🚀 Quick Start

### Prerequisites
- Windows 10/11 or Windows Server 2016+
- .NET 6.0 Runtime (download from https://dotnet.microsoft.com/download)
- Domain access with appropriate permissions

### Installation Options

#### Option 1: Run Pre-built Executable
1. Download the latest release
2. Extract the ZIP file
3. Run `ADAUDIT.exe`

#### Option 2: Build from Source
1. **Using PowerShell (Recommended)**:
   ```powershell
   .\build.ps1
   ```

2. **Using Command Line**:
   ```cmd
   .\build.bat
   ```

3. **Using Visual Studio**:
   - Open `ADAUDIT.sln`
   - Press `Ctrl+Shift+B` to build
   - Press `F5` to run

### First Run
1. Launch the application
2. Go to the **Connection** tab
3. Enter your domain controller details:
   - **Domain Controller**: `your-dc-server.domain.com`
   - **Username**: `domain\username` or `username`
   - **Password**: Your domain password
4. Click **Connect**
5. Once connected, use the other tabs to explore users, groups, and GPOs

## 🔧 Configuration

### Connection Settings
- **Default LDAP Port**: 389 (standard) / 636 (SSL)
- **Connection Timeout**: 30 seconds
- **Search Timeout**: 30 seconds

### Security Settings
- **SSL Required**: Disabled by default (enable for production)
- **Password Complexity**: Enabled
- **Minimum Password Length**: 8 characters

## 📋 Features Overview

### Users Tab
- View all domain users
- See account status (enabled/disabled, locked)
- Change user passwords (with proper permissions)
- View last logon times and account details

### Groups Tab
- List all domain groups
- View group types and scopes
- See member counts and details
- Analyze group memberships

### Group Policies Tab
- List all GPOs in the domain
- View GPO status (enabled/disabled)
- See GPO versions and details
- Analyze GPO configurations

## 🛠️ Troubleshooting

### Common Issues

**"Connection failed"**
- Check domain controller name/IP
- Verify network connectivity
- Ensure credentials are correct

**"Access denied"**
- Check user permissions in Active Directory
- Verify account is not locked/expired
- Ensure proper domain membership

**"No users/groups found"**
- Check LDAP search permissions
- Verify connection to correct domain
- Review firewall settings

### Performance Tips
- Use local domain controllers when possible
- For large domains, consider filtering by OU
- Close unused tabs to reduce memory usage

## 📞 Support

For issues:
1. Check the troubleshooting section in README.md
2. Verify your environment meets requirements
3. Test with a simple domain account first
4. Review application logs if available

## 🔒 Security Notes

- Always use SSL (LDAPS) in production environments
- Use dedicated service accounts with minimal permissions
- Regularly audit access and permissions
- Follow your organization's security policies

---

**Ready to start?** Run the application and connect to your domain! 