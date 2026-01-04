# PowerShell Scripts Collection

A comprehensive collection of PowerShell scripts for Active Directory management, Exchange administration, system monitoring, and security operations. All scripts are production-tested and include detailed comments.

## 📁 Repository Structure

```
Powershell/
├── Active-Directory/          # AD user, group, and OU management
├── Exchange/                  # Exchange Server administration
├── Monitoring-Security/       # System monitoring and security tools
├── Access-Control/            # SMB shares and permissions
├── Utilities/                 # General utility scripts
└── CommonADUserTasks/         # Common AD task collection
```

## 📚 Script Categories

### 🔐 Active Directory Management
Scripts for managing users, groups, OUs, and AD objects.

| Script | Description |
|--------|-------------|
| [GenerateUsersFromCSV_WithOutPromptUser.ps1](./Active-Directory/GenerateUsersFromCSV_WithOutPromptUser.ps1) | Comprehensive user provisioning for AD and Exchange environments from CSV |
| [ExtractADgroupMembers.ps1](./Active-Directory/ExtractADgroupMembers.ps1) | Extract and analyze AD group memberships with detailed reporting |
| [Disabled_Expiring_MoveUser.ps1](./Active-Directory/Disabled_Expiring_MoveUser.ps1) | Disable expired users and manage organizational units (OUs) |
| [logontimes_Computers.ps1](./Active-Directory/logontimes_Computers.ps1) | Monitor computer login times across Active Directory domains |
| [DoesITExsists_Files_Folder.ps1](./Active-Directory/DoesITExsists_Files_Folder.ps1) | Search files across all AD computers with pattern matching |
| [ResetUserPassword_byList.ps1](./Active-Directory/ResetUserPassword_byList.ps1) | Bulk password reset from user list with secure password generation |

### 📧 Exchange Administration
Scripts for Exchange Server mailbox and user management.

| Script | Description |
|--------|-------------|
| [ExchangeHack01.ps1](./Exchange/ExchangeHack01.ps1) | Access and view user mailbox data, permissions, and properties |

### 🔍 Monitoring & Security
System monitoring, security detection, and inventory tools.

| Script | Description |
|--------|-------------|
| [CompareProcessesRunning.ps1](./Monitoring-Security/CompareProcessesRunning.ps1) | Detect unauthorized processes and potential backdoors by comparing running processes |
| [FindServicesOnComputer.ps1](./Monitoring-Security/FindServicesOnComputer.ps1) | Interactive service discovery and management across remote computers |
| [UptimeComputerAndLoggedInUsers.ps1](./Monitoring-Security/UptimeComputerAndLoggedInUsers.ps1) | Monitor system uptime and currently logged-in users |
| [testJavaversions.ps1](./Monitoring-Security/testJavaversions.ps1) | Detect Java versions installed on remote computers for security auditing |

### 🔒 Access Control & Permissions
Scripts for managing file shares, SMB permissions, and access control.

| Script | Description |
|--------|-------------|
| [PowershellGetShares_WhoHasAccess.ps1](./Access-Control/PowershellGetShares_WhoHasAccess.ps1) | Manage and configure SMB share permissions with detailed access reporting |

### 🛠️ Utilities
General-purpose utility scripts.

| Script | Description |
|--------|-------------|
| [randpasswd.ps1](./Utilities/randpasswd.ps1) | Generate secure random passwords with customizable length and complexity |

### 📦 Common AD Tasks Collection
A collection of essential AD management scripts in the `CommonADUserTasks/` folder:

| Script | Description |
|--------|-------------|
| [ChangeUsersProperties_FromList.ps1](./CommonADUserTasks/ChangeUsersProperties_FromList.ps1) | Bulk update user properties from CSV list |
| [CommonADTasksWithUserList.ps1](./CommonADUserTasks/CommonADTasksWithUserList.ps1) | Common AD tasks with user list processing |
| [FindMemberOfGroupsFromUser.ps1](./CommonADUserTasks/FindMemberOfGroupsFromUser.ps1) | Find all group memberships for specified users |
| [GroupMembers.ps1](./CommonADUserTasks/GroupMembers.ps1) | Extract and manage group members with export capabilities |
| [UsersPasswordLastSet_ToCSV.ps1](./CommonADUserTasks/UsersPasswordLastSet_ToCSV.ps1) | Export password last set dates to CSV for compliance reporting |

## 🚀 Quick Start

### Prerequisites

- **PowerShell 5.1 or later** (Windows PowerShell or PowerShell Core)
- **Active Directory PowerShell Module** (for AD scripts)
  ```powershell
  Install-WindowsFeature RSAT-AD-PowerShell
  ```
- **Exchange Management Shell** (for Exchange scripts)
- **Appropriate permissions** for the operations being performed

### Usage Example

```powershell
# Example: Generate users from CSV
.\Active-Directory\GenerateUsersFromCSV_WithOutPromptUser.ps1 -CSVPath "users.csv"

# Example: Extract AD group members
.\Active-Directory\ExtractADgroupMembers.ps1 -GroupName "Domain Admins"

# Example: Monitor computer uptime
.\Monitoring-Security\UptimeComputerAndLoggedInUsers.ps1 -ComputerName "SERVER01"
```

## 📖 Documentation

Each script includes:
- **Header comments** explaining purpose and usage
- **Parameter descriptions** with examples
- **Error handling** and logging
- **Output formatting** for readability

## ⚠️ Security Notice

**IMPORTANT:** These scripts are provided as-is for educational and administrative purposes.

### Before Running:
- ✅ Review scripts before execution
- ✅ Test in a non-production environment first
- ✅ Ensure you have proper authorization
- ✅ Follow your organization's change management procedures
- ✅ Backup critical data before making changes
- ✅ Verify script behavior matches your requirements

### Security Best Practices:
- Run scripts with least privilege required
- Audit script execution in production environments
- Review and sanitize any CSV input files
- Use secure password generation for bulk operations
- Monitor script execution logs

## 🔗 Related Projects

- **[Portfolio Website](https://maxjensen.dk)** - View these scripts in action on my portfolio
- **[C# Projects](https://github.com/R3verse)** - Related C# applications for user management

## 📊 Statistics

- **Total Scripts:** 18+ individual scripts
- **Categories:** 6 organized folders
- **Languages:** PowerShell 5.1+
- **License:** Educational and administrative use

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to:
- Open an issue for bugs or feature requests
- Submit a pull request with improvements
- Share your use cases and feedback

## 📝 License

This collection is provided for educational and administrative purposes. Use at your own risk. Always review and test scripts before using in production environments.

## 📧 Contact

For questions, suggestions, or contributions:
- **Portfolio:** [maxjensen.dk](https://maxjensen.dk)
- **GitHub:** [@R3verse](https://github.com/R3verse)

---

**Last Updated:** January 2026  
**Maintained by:** Max Jensen
