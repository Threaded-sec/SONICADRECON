using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using ADAUDIT.Services;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Drawing.Drawing2D;

namespace ADAUDIT
{
    public partial class MainForm : Form
    {
        private readonly LDAPService _ldapService;
        private readonly UserManagementService _userService;
        private readonly GroupManagementService _groupService;
        private readonly GroupPolicyService _gpoService;
        private readonly AuditService _auditService;

        public MainForm()
        {
            InitializeComponent();
            
            // Set form icon
            try
            {
                if (System.IO.File.Exists("app.ico"))
                {
                    this.Icon = new System.Drawing.Icon("app.ico");
                }
            }
            catch
            {
                // Continue without icon if loading fails
            }
            
            _ldapService = new LDAPService();
            _userService = new UserManagementService(_ldapService);
            _groupService = new GroupManagementService(_ldapService);
            _gpoService = new GroupPolicyService(_ldapService);
            _auditService = new AuditService(_ldapService);
            
            InitializeConnectionPanel();
            InitializeUserPanel();
            InitializeGroupPanel();
            InitializeGPOPanel();
            InitializeAuditPanel();
            InitializeLogoutButton();
            
            // Disable functional tabs until connection is established
            for (int i = 1; i < tabControl.TabPages.Count; i++)
            {
                tabControl.TabPages[i].Enabled = false;
            }
        }

        /// <summary>
        /// Adds watermark to PDF document
        /// </summary>
        private void AddWatermarkToPDF(Document document, PdfWriter writer)
        {
            try
            {
                // Create watermark text with transparency
                var watermarkFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 36);
                var watermark = new Chunk("SONIC-AUDIT", watermarkFont);
                watermark.Font.Color = BaseColor.LIGHT_GRAY;
                
                // Create watermark phrase
                var watermarkPhrase = new Phrase(watermark);
                
                // Add watermark to the document
                var watermarkParagraph = new Paragraph(watermarkPhrase);
                watermarkParagraph.Alignment = Element.ALIGN_CENTER;
                watermarkParagraph.SpacingBefore = 200f;
                watermarkParagraph.SpacingAfter = 200f;
                
                document.Add(watermarkParagraph);
            }
            catch
            {
                // Continue without watermark if it fails
            }
        }

        /// <summary>
        /// Adds watermark to Excel worksheet
        /// </summary>
        private void AddWatermarkToExcel(OfficeOpenXml.ExcelWorksheet worksheet)
        {
            try
            {
                // Add watermark text to header
                var watermarkCell = worksheet.Cells[1, 1];
                watermarkCell.Value = "SONIC-AUDIT";
                watermarkCell.Style.Font.Size = 14;
                watermarkCell.Style.Font.Bold = true;
                watermarkCell.Style.Font.Color.SetColor(System.Drawing.Color.LightGray);
                
                // Merge cells for watermark if needed
                if (worksheet.Dimension != null)
                {
                    var watermarkRange = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
                    watermarkRange.Merge = true;
                    watermarkRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }
            }
            catch
            {
                // Continue without watermark if it fails
            }
        }

        private void InitializeConnectionPanel()
        {
            // Connection Tab
            var connectionTab = new TabPage("Connection");
            connectionTab.BackColor = Color.FromArgb(240, 240, 245);

            var connectionPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20)
            };

            // Center the form elements
            int centerX = 500; // Center of the 1000px wide window
            int startY = 80;
            int labelWidth = 150;
            int textBoxWidth = 250;
            int spacing = 50;

            var lblTitle = new Label 
            { 
                Text = "SONIC AD RECON", 
                Location = new Point(centerX - 80, startY - 80), 
                AutoSize = true, 
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Bold),
                ForeColor = Color.FromArgb(50, 50, 50),
                TextAlign = ContentAlignment.MiddleCenter
            };

            var lblServer = new Label 
            { 
                Text = "Domain Controller:", 
                Location = new Point(centerX - labelWidth - 10, startY + 20), 
                AutoSize = true,
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular),
                ForeColor = Color.FromArgb(70, 70, 70)
            };
            var txtServer = new TextBox 
            { 
                Name = "txtServer", 
                Location = new Point(centerX + 30, startY + 17), 
                Width = textBoxWidth, 
                Text = "localhost", 
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 10),
                BorderStyle = BorderStyle.FixedSingle
            };

            var lblUsername = new Label 
            { 
                Text = "Username:", 
                Location = new Point(centerX - labelWidth - 10, startY + spacing + 20), 
                AutoSize = true,
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular),
                ForeColor = Color.FromArgb(70, 70, 70)
            };
            var txtUsername = new TextBox 
            { 
                Name = "txtUsername", 
                Location = new Point(centerX + 30, startY + spacing + 17), 
                Width = textBoxWidth,
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 10),
                BorderStyle = BorderStyle.FixedSingle
            };

            var lblPassword = new Label 
            { 
                Text = "Password:", 
                Location = new Point(centerX - labelWidth - 10, startY + spacing * 2 + 20), 
                AutoSize = true,
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 10, FontStyle.Regular),
                ForeColor = Color.FromArgb(70, 70, 70)
            };
            var txtPassword = new TextBox 
            { 
                Name = "txtPassword", 
                Location = new Point(centerX + 30, startY + spacing * 2 + 17), 
                Width = textBoxWidth, 
                PasswordChar = '*',
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 10),
                BorderStyle = BorderStyle.FixedSingle
            };

            var btnConnect = new Button
            {
                Text = "Connect",
                Location = new Point(centerX - 60, startY + spacing * 3 + 20),
                Width = 120,
                Height = 35,
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 10, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            var lblStatus = new Label 
            { 
                Name = "lblStatus", 
                Location = new Point(centerX - 100, startY + spacing * 4 + 20), 
                AutoSize = true, 
                ForeColor = Color.Gray,
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 9),
                TextAlign = ContentAlignment.MiddleCenter
            };

            btnConnect.Click += (sender, e) =>
            {
                try
                {
                    _ldapService.Connect(txtServer.Text, txtUsername.Text, txtPassword.Text);
                    lblStatus.Text = "Connected successfully!";
                    lblStatus.ForeColor = Color.Green;
                    
                    // Enable the functional tabs after successful connection
                    for (int i = 1; i < tabControl.TabPages.Count; i++)
                    {
                        tabControl.TabPages[i].Enabled = true;
                    }
                    
                    // Show logout button
                    var logoutButton = this.Controls.OfType<Button>().FirstOrDefault(b => b.Text == "Logout");
                    if (logoutButton != null)
                    {
                        logoutButton.Visible = true;
                    }
                    
                    // Automatically switch to Users tab after successful connection
                    tabControl.SelectedIndex = 1; // Users tab is at index 1
                }
                catch (Exception ex)
                {
                    lblStatus.Text = $"Connection failed: {ex.Message}";
                    lblStatus.ForeColor = Color.Red;
                    
                    // Disable functional tabs on connection failure
                    for (int i = 1; i < tabControl.TabPages.Count; i++)
                    {
                        tabControl.TabPages[i].Enabled = false;
                    }
                }
            };

            connectionPanel.Controls.AddRange(new Control[] { lblTitle, lblServer, txtServer, lblUsername, txtUsername, lblPassword, txtPassword, btnConnect, lblStatus });
            connectionTab.Controls.Add(connectionPanel);
            tabControl.TabPages.Add(connectionTab);
        }

        private void InitializeUserPanel()
        {
            // Users Tab
            var usersTab = new TabPage("Users");
            usersTab.BackColor = Color.White;

            var usersPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            var btnRefreshUsers = new Button
            {
                Text = "Refresh Users",
                Location = new Point(10, 10),
                Width = 120,
                Height = 30
            };

            var btnExportUsers = new Button
            {
                Text = "XLS",
                Location = new Point(140, 10),
                Width = 120,
                Height = 30
            };

            var btnExportUsersPDF = new Button
            {
                Text = "PDF",
                Location = new Point(270, 10),
                Width = 120,
                Height = 30
            };

            var usersListView = new ListView
            {
                Name = "usersListView",
                Location = new Point(10, 50),
                Size = new Size(760, 350),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                MultiSelect = false
            };

            usersListView.Columns.Add("Username", 180);
            usersListView.Columns.Add("Display Name", 250);
            usersListView.Columns.Add("Email", 250);
            usersListView.Columns.Add("Last Logon", 180);

            var userDetailsPanel = new Panel
            {
                Location = new Point(10, 410),
                Size = new Size(760, 180),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            var lblSelectedUser = new Label { Text = "Selected User:", Location = new Point(10, 10), AutoSize = true };
            var txtSelectedUser = new TextBox { Name = "txtSelectedUser", Location = new Point(120, 7), Width = 200, ReadOnly = true };

            var lblNewPassword = new Label { Text = "New Password:", Location = new Point(10, 40), AutoSize = true };
            var txtNewPassword = new TextBox { Name = "txtNewPassword", Location = new Point(120, 37), Width = 200, PasswordChar = '*' };

            var lblConfirmPassword = new Label { Text = "Confirm Password:", Location = new Point(10, 70), AutoSize = true };
            var txtConfirmPassword = new TextBox { Name = "txtConfirmPassword", Location = new Point(120, 67), Width = 200, PasswordChar = '*' };

            var btnChangePassword = new Button
            {
                Text = "Change Password",
                Location = new Point(120, 100),
                Width = 120,
                Height = 30,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            var btnClearPassword = new Button
            {
                Text = "Clear",
                Location = new Point(250, 100),
                Width = 80,
                Height = 30,
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };

            btnRefreshUsers.Click += (sender, e) => RefreshUsers(usersListView);
            btnExportUsers.Click += (sender, e) => ExportUsersToExcel(usersListView);
            btnExportUsersPDF.Click += (sender, e) => ExportUsersToPDF(usersListView);
            btnChangePassword.Click += (sender, e) => ChangeUserPassword(txtSelectedUser.Text, txtNewPassword.Text, txtConfirmPassword.Text);
            btnClearPassword.Click += (sender, e) => 
            {
                txtNewPassword.Text = string.Empty;
                txtConfirmPassword.Text = string.Empty;
            };

            usersListView.SelectedIndexChanged += (sender, e) =>
            {
                if (usersListView.SelectedItems.Count > 0)
                {
                    txtSelectedUser.Text = usersListView.SelectedItems[0].Text;
                }
            };

            userDetailsPanel.Controls.AddRange(new Control[] { lblSelectedUser, txtSelectedUser, lblNewPassword, txtNewPassword, lblConfirmPassword, txtConfirmPassword, btnChangePassword, btnClearPassword });
            usersPanel.Controls.AddRange(new Control[] { btnRefreshUsers, btnExportUsers, btnExportUsersPDF, usersListView, userDetailsPanel });
            usersTab.Controls.Add(usersPanel);
            tabControl.TabPages.Add(usersTab);
        }

        private void InitializeGroupPanel()
        {
            // Groups Tab
            var groupsTab = new TabPage("Groups");
            groupsTab.BackColor = Color.White;

            var groupsPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            var btnRefreshGroups = new Button
            {
                Text = "Refresh Groups",
                Location = new Point(10, 10),
                Width = 120,
                Height = 30
            };

            var btnExportGroups = new Button
            {
                Text = "XLS",
                Location = new Point(140, 10),
                Width = 120,
                Height = 30
            };

            var groupsListView = new ListView
            {
                Name = "groupsListView",
                Location = new Point(10, 50),
                Size = new Size(760, 300),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            groupsListView.Columns.Add("Group Name", 250);
            groupsListView.Columns.Add("Description", 400);
            groupsListView.Columns.Add("Members Count", 150);

            btnRefreshGroups.Click += (sender, e) => RefreshGroups(groupsListView);
            btnExportGroups.Click += (sender, e) => ExportGroupsToExcel(groupsListView);

            groupsPanel.Controls.AddRange(new Control[] { btnRefreshGroups, btnExportGroups, groupsListView });
            groupsTab.Controls.Add(groupsPanel);
            tabControl.TabPages.Add(groupsTab);
        }

        private void InitializeGPOPanel()
        {
            // Group Policies Tab
            var gpoTab = new TabPage("Group Policies");
            gpoTab.BackColor = Color.White;

            var gpoPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            var btnRefreshGPOs = new Button
            {
                Text = "Refresh GPOs",
                Location = new Point(10, 10),
                Width = 120,
                Height = 30
            };

            var btnExportGPOs = new Button
            {
                Text = "XLS",
                Location = new Point(140, 10),
                Width = 120,
                Height = 30
            };

            var gpoListView = new ListView
            {
                Name = "gpoListView",
                Location = new Point(10, 50),
                Size = new Size(760, 300),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            gpoListView.Columns.Add("GPO Name", 250);
            gpoListView.Columns.Add("ID", 400);
            gpoListView.Columns.Add("Status", 150);

            btnRefreshGPOs.Click += (sender, e) => RefreshGPOs(gpoListView);
            btnExportGPOs.Click += (sender, e) => ExportGPOsToExcel(gpoListView);

            gpoPanel.Controls.AddRange(new Control[] { btnRefreshGPOs, btnExportGPOs, gpoListView });
            gpoTab.Controls.Add(gpoPanel);
            tabControl.TabPages.Add(gpoTab);
        }

        private void InitializeLogoutButton()
        {
            // Create logout button in the main form
            var btnLogout = new Button
            {
                Text = "Logout",
                Location = new Point(this.ClientSize.Width - 100, this.ClientSize.Height - 50),
                Width = 80,
                Height = 30,
                Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 9, FontStyle.Bold),
                BackColor = Color.FromArgb(220, 53, 69), // Red color for logout
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Visible = false, // Initially hidden until connected
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right // Anchor to bottom right
            };

            btnLogout.Click += (sender, e) => Logout();

            // Add logout button to the main form
            this.Controls.Add(btnLogout);
            btnLogout.BringToFront();
        }

        private void Logout()
        {
            try
            {
                // Disconnect from LDAP
                _ldapService.Disconnect();
                
                // Disable all functional tabs
                for (int i = 1; i < tabControl.TabPages.Count; i++)
                {
                    tabControl.TabPages[i].Enabled = false;
                }
                
                // Switch back to connection tab
                tabControl.SelectedIndex = 0;
                
                // Clear all data from tabs
                ClearAllData();
                
                // Hide logout button
                var logoutButton = this.Controls.OfType<Button>().FirstOrDefault(b => b.Text == "Logout");
                if (logoutButton != null)
                {
                    logoutButton.Visible = false;
                }
                
                // Update connection status
                var statusLabel = FindControlRecursive(this, "lblStatus") as Label;
                if (statusLabel != null)
                {
                    statusLabel.Text = "Disconnected";
                    statusLabel.ForeColor = Color.Gray;
                }
                
                MessageBox.Show("Successfully logged out. Please reconnect to continue.", "Logout", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during logout: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearAllData()
        {
            // Clear Users tab
            var usersListView = FindControlRecursive(this, "usersListView") as ListView;
            if (usersListView != null)
            {
                usersListView.Items.Clear();
            }

            // Clear Groups tab
            var groupsListView = FindControlRecursive(this, "groupsListView") as ListView;
            if (groupsListView != null)
            {
                groupsListView.Items.Clear();
            }

            // Clear GPOs tab
            var gpoListView = FindControlRecursive(this, "gpoListView") as ListView;
            if (gpoListView != null)
            {
                gpoListView.Items.Clear();
            }

            // Clear Audit tab
            var summaryListView = FindControlRecursive(this, "summaryListView") as ListView;
            if (summaryListView != null)
            {
                summaryListView.Items.Clear();
            }

            var auditListView = FindControlRecursive(this, "auditListView") as ListView;
            if (auditListView != null)
            {
                auditListView.Items.Clear();
            }
        }

        private Control FindControlRecursive(Control parent, string controlName)
        {
            if (parent.Name == controlName)
                return parent;

            foreach (Control child in parent.Controls)
            {
                var found = FindControlRecursive(child, controlName);
                if (found != null)
                    return found;
            }

            return null;
        }

        private void InitializeAuditPanel()
        {
            // Audit Tab
            var auditTab = new TabPage("Security Audit");
            auditTab.BackColor = Color.White;

            var auditPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            // Summary Panel
            var summaryPanel = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(760, 180),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var lblSummary = new Label { Text = "Security Summary:", Location = new Point(10, 10), AutoSize = true, Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 10, FontStyle.Bold) };
            var summaryListView = new ListView
            {
                Name = "summaryListView",
                Location = new Point(10, 30),
                Size = new Size(740, 140),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            summaryListView.Columns.Add("Category", 200);
            summaryListView.Columns.Add("Count", 100);
            summaryListView.Columns.Add("Risk Level", 100);

            // Audit Buttons Panel
            var auditButtonsPanel = new Panel
            {
                Location = new Point(10, 200),
                Size = new Size(760, 50),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            var btnRunFullAudit = new Button
            {
                Text = "Run Full Security Audit",
                Location = new Point(10, 10),
                Width = 150,
                Height = 30
            };

            var btnComprehensiveAudit = new Button
            {
                Text = "Comprehensive Audit",
                Location = new Point(170, 10),
                Width = 150,
                Height = 30
            };

            var btnWeakPasswords = new Button
            {
                Text = "Weak Passwords",
                Location = new Point(330, 10),
                Width = 120,
                Height = 30
            };

            var btnPrivilegedAccounts = new Button
            {
                Text = "Privileged Accounts",
                Location = new Point(460, 10),
                Width = 120,
                Height = 30
            };

            var btnInactiveAccounts = new Button
            {
                Text = "Inactive Accounts",
                Location = new Point(590, 10),
                Width = 120,
                Height = 30
            };

            var btnLockedAccounts = new Button
            {
                Text = "Locked Accounts",
                Location = new Point(720, 10),
                Width = 120,
                Height = 30
            };

            var btnExportAudit = new Button
            {
                Text = "XLS",
                Location = new Point(850, 10),
                Width = 120,
                Height = 30
            };

            var btnExportAuditPDF = new Button
            {
                Text = "PDF",
                Location = new Point(980, 10),
                Width = 120,
                Height = 30
            };

            // Detailed Audit Results
            var auditListView = new ListView
            {
                Name = "auditListView",
                Location = new Point(10, 260),
                Size = new Size(760, 240),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            auditListView.Columns.Add("Username", 120);
            auditListView.Columns.Add("Display Name", 150);
            auditListView.Columns.Add("Security Issues", 300);
            auditListView.Columns.Add("Password Age", 100);
            auditListView.Columns.Add("Last Logon", 120);
            auditListView.Columns.Add("Account Status", 100);

            // Event handlers
            btnRunFullAudit.Click += (sender, e) => RunFullSecurityAudit(summaryListView, auditListView);
            btnComprehensiveAudit.Click += (sender, e) => RunComprehensiveSecurityAudit(summaryListView, auditListView);
            btnWeakPasswords.Click += (sender, e) => ShowWeakPasswords(auditListView);
            btnPrivilegedAccounts.Click += (sender, e) => ShowPrivilegedAccounts(auditListView);
            btnInactiveAccounts.Click += (sender, e) => ShowInactiveAccounts(auditListView);
            btnLockedAccounts.Click += (sender, e) => ShowLockedAccounts(auditListView);
            btnExportAudit.Click += (sender, e) => ExportAuditReport(auditListView);
            btnExportAuditPDF.Click += (sender, e) => ExportAuditReportToPDF(auditListView);

            summaryPanel.Controls.AddRange(new Control[] { lblSummary, summaryListView });
            auditButtonsPanel.Controls.AddRange(new Control[] { btnRunFullAudit, btnComprehensiveAudit, btnWeakPasswords, btnPrivilegedAccounts, btnInactiveAccounts, btnLockedAccounts, btnExportAudit, btnExportAuditPDF });
            auditPanel.Controls.AddRange(new Control[] { summaryPanel, auditButtonsPanel, auditListView });
            auditTab.Controls.Add(auditPanel);
            tabControl.TabPages.Add(auditTab);
        }

        private void RefreshUsers(ListView listView)
        {
            try
            {
                listView.Items.Clear();
                var users = _userService.GetAllUsers();
                
                foreach (var user in users)
                {
                    var item = new ListViewItem(user.Username);
                    item.SubItems.Add(user.DisplayName);
                    item.SubItems.Add(user.Email);
                    item.SubItems.Add(user.LastLogon?.ToString("yyyy-MM-dd HH:mm") ?? "Never");
                    listView.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing users: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshGroups(ListView listView)
        {
            try
            {
                listView.Items.Clear();
                var groups = _groupService.GetAllGroups();
                
                foreach (var group in groups)
                {
                    var item = new ListViewItem(group.Name);
                    item.SubItems.Add(group.Description);
                    item.SubItems.Add(group.MemberCount.ToString());
                    listView.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing groups: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshGPOs(ListView listView)
        {
            try
            {
                listView.Items.Clear();
                var gpos = _gpoService.GetAllGPOs();
                
                foreach (var gpo in gpos)
                {
                    var item = new ListViewItem(gpo.Name);
                    item.SubItems.Add(gpo.Id);
                    item.SubItems.Add(gpo.Status);
                    listView.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing GPOs: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ChangeUserPassword(string username, string newPassword, string confirmPassword)
        {
            if (string.IsNullOrEmpty(username))
            {
                MessageBox.Show("Please select a user first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(newPassword))
            {
                MessageBox.Show("Please enter a new password.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(confirmPassword))
            {
                MessageBox.Show("Please confirm the new password.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (newPassword != confirmPassword)
            {
                MessageBox.Show("Passwords do not match. Please ensure both passwords are identical.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (newPassword.Length < 8)
            {
                var result = MessageBox.Show("Password is less than 8 characters. This may not meet your domain's password policy. Continue anyway?", "Password Policy Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result != DialogResult.Yes)
                {
                    return;
                }
            }

            try
            {
                _userService.ChangePassword(username, newPassword);
                MessageBox.Show($"Password changed successfully for user: {username}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // Clear password fields after successful change
                var txtNewPassword = FindControlRecursive(this, "txtNewPassword") as TextBox;
                var txtConfirmPassword = FindControlRecursive(this, "txtConfirmPassword") as TextBox;
                if (txtNewPassword != null) txtNewPassword.Text = string.Empty;
                if (txtConfirmPassword != null) txtConfirmPassword.Text = string.Empty;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing password: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportUsersToExcel(ListView listView)
        {
            try
            {
                if (listView.Items.Count == 0)
                {
                    MessageBox.Show("No data to export. Please refresh the users list first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = $"AD_Users_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using var package = new OfficeOpenXml.ExcelPackage();
                    var worksheet = package.Workbook.Worksheets.Add("Users");

                    // Add watermark
                    AddWatermarkToExcel(worksheet);

                    // Add headers (shifted down by 1 row to make room for watermark)
                    for (int i = 0; i < listView.Columns.Count; i++)
                    {
                        worksheet.Cells[2, i + 1].Value = listView.Columns[i].Text;
                        worksheet.Cells[2, i + 1].Style.Font.Bold = true;
                    }

                    // Add data (shifted down by 1 row to account for watermark)
                    for (int row = 0; row < listView.Items.Count; row++)
                    {
                        var item = listView.Items[row];
                        for (int col = 0; col < item.SubItems.Count; col++)
                        {
                            worksheet.Cells[row + 3, col + 1].Value = item.SubItems[col].Text;
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells.AutoFitColumns();

                    package.SaveAs(new FileInfo(saveFileDialog.FileName));
                    MessageBox.Show($"Users exported successfully to: {saveFileDialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting users: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportGroupsToExcel(ListView listView)
        {
            try
            {
                if (listView.Items.Count == 0)
                {
                    MessageBox.Show("No data to export. Please refresh the groups list first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = $"AD_Groups_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using var package = new OfficeOpenXml.ExcelPackage();
                    var worksheet = package.Workbook.Worksheets.Add("Groups");

                    // Add watermark
                    AddWatermarkToExcel(worksheet);

                    // Add headers (shifted down by 1 row to make room for watermark)
                    for (int i = 0; i < listView.Columns.Count; i++)
                    {
                        worksheet.Cells[2, i + 1].Value = listView.Columns[i].Text;
                        worksheet.Cells[2, i + 1].Style.Font.Bold = true;
                    }

                    // Add data (shifted down by 1 row to account for watermark)
                    for (int row = 0; row < listView.Items.Count; row++)
                    {
                        var item = listView.Items[row];
                        for (int col = 0; col < item.SubItems.Count; col++)
                        {
                            worksheet.Cells[row + 3, col + 1].Value = item.SubItems[col].Text;
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells.AutoFitColumns();

                    package.SaveAs(new FileInfo(saveFileDialog.FileName));
                    MessageBox.Show($"Groups exported successfully to: {saveFileDialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting groups: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportGPOsToExcel(ListView listView)
        {
            try
            {
                if (listView.Items.Count == 0)
                {
                    MessageBox.Show("No data to export. Please refresh the GPOs list first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = $"AD_GPOs_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using var package = new OfficeOpenXml.ExcelPackage();
                    var worksheet = package.Workbook.Worksheets.Add("Group Policies");

                    // Add watermark
                    AddWatermarkToExcel(worksheet);

                    // Add headers (shifted down by 1 row to make room for watermark)
                    for (int i = 0; i < listView.Columns.Count; i++)
                    {
                        worksheet.Cells[2, i + 1].Value = listView.Columns[i].Text;
                        worksheet.Cells[2, i + 1].Style.Font.Bold = true;
                    }

                    // Add data (shifted down by 1 row to account for watermark)
                    for (int row = 0; row < listView.Items.Count; row++)
                    {
                        var item = listView.Items[row];
                        for (int col = 0; col < item.SubItems.Count; col++)
                        {
                            worksheet.Cells[row + 3, col + 1].Value = item.SubItems[col].Text;
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells.AutoFitColumns();

                    package.SaveAs(new FileInfo(saveFileDialog.FileName));
                    MessageBox.Show($"Group Policies exported successfully to: {saveFileDialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting GPOs: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportUsersToPDF(ListView listView)
        {
            try
            {
                if (listView.Items.Count == 0)
                {
                    MessageBox.Show("No data to export. Please refresh the users list first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using var saveFileDialog = new SaveFileDialog
                {
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    FileName = $"AD_Users_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using var fileStream = new FileStream(saveFileDialog.FileName, FileMode.Create);
                    using var document = new Document(PageSize.A4, 25, 25, 30, 30);
                    var writer = PdfWriter.GetInstance(document, fileStream);

                    document.Open();

                    // Add title
                    var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18);
                    var title = new Paragraph("Active Directory Users Report", titleFont);
                    title.Alignment = Element.ALIGN_CENTER;
                    title.SpacingAfter = 20f;
                    document.Add(title);

                    // Add generation date
                    var dateFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);
                    var dateText = new Paragraph($"Generated on: {DateTime.Now:yyyy-MM-dd HH:mm:ss}", dateFont);
                    dateText.Alignment = Element.ALIGN_CENTER;
                    dateText.SpacingAfter = 20f;
                    document.Add(dateText);

                    // Add summary section
                    var summaryFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14);
                    var summaryTitle = new Paragraph("Users Summary", summaryFont);
                    summaryTitle.SpacingAfter = 10f;
                    document.Add(summaryTitle);

                    var summaryText = new Paragraph($"Total users: {listView.Items.Count}", dateFont);
                    summaryText.SpacingAfter = 15f;
                    document.Add(summaryText);

                    // Create table for users
                    var table = new PdfPTable(listView.Columns.Count);
                    table.WidthPercentage = 100;

                    // Add headers
                    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                    foreach (ColumnHeader column in listView.Columns)
                    {
                        var cell = new PdfPCell(new Phrase(column.Text, headerFont));
                        cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                        cell.Padding = 5f;
                        table.AddCell(cell);
                    }

                    // Add data rows
                    var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);
                    foreach (ListViewItem item in listView.Items)
                    {
                        foreach (ListViewItem.ListViewSubItem subItem in item.SubItems)
                        {
                            var cell = new PdfPCell(new Phrase(subItem.Text, dataFont));
                            cell.Padding = 3f;
                            table.AddCell(cell);
                        }
                    }

                    document.Add(table);

                    // Add watermark
                    AddWatermarkToPDF(document, writer);

                    // Add footer
                    document.Add(new Paragraph(" ")); // Spacing
                    var footerFont = FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8);
                    var footer = new Paragraph("This report was generated by SonicADRecon", footerFont);
                    footer.Alignment = Element.ALIGN_CENTER;
                    document.Add(footer);

                    document.Close();
                    writer.Close();

                    MessageBox.Show($"Users report exported successfully to: {saveFileDialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting users to PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Audit Methods
        private void RunFullSecurityAudit(ListView summaryListView, ListView auditListView)
        {
            try
            {
                summaryListView.Items.Clear();
                auditListView.Items.Clear();

                var summary = _auditService.GetSecurityIssueSummary();
                var users = _auditService.GetUsersWithSecurityIssues();

                // Populate summary
                foreach (var item in summary)
                {
                    var listItem = new ListViewItem(item.Key);
                    listItem.SubItems.Add(item.Value.ToString());
                    
                    // Determine risk level
                    string riskLevel = "Low";
                    if (item.Key.Contains("Weak Passwords") || item.Key.Contains("Privileged Accounts"))
                        riskLevel = "High";
                    else if (item.Key.Contains("Locked") || item.Key.Contains("Inactive"))
                        riskLevel = "Medium";
                    
                    listItem.SubItems.Add(riskLevel);
                    summaryListView.Items.Add(listItem);
                }

                // Populate detailed results
                foreach (var user in users.Where(u => u.SecurityIssues.Count > 0))
                {
                    var item = new ListViewItem(user.Username);
                    item.SubItems.Add(user.DisplayName);
                    item.SubItems.Add(string.Join("; ", user.SecurityIssues));
                    item.SubItems.Add(user.PasswordAge?.ToString() ?? "N/A");
                    item.SubItems.Add(user.LastLogon?.ToString("yyyy-MM-dd") ?? "Never");
                    
                    string status = user.IsEnabled ? "Enabled" : "Disabled";
                    if (user.AccountLocked) status += " (Locked)";
                    item.SubItems.Add(status);
                    
                    auditListView.Items.Add(item);
                }

                MessageBox.Show($"Security audit completed. Found {users.Count(u => u.SecurityIssues.Count > 0)} users with security issues.", 
                    "Audit Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error running security audit: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RunComprehensiveSecurityAudit(ListView summaryListView, ListView auditListView)
        {
            try
            {
                summaryListView.Items.Clear();
                auditListView.Items.Clear();

                var comprehensiveResults = _auditService.RunComprehensiveSecurityAudit();

                // Populate summary with comprehensive audit results
                foreach (var result in comprehensiveResults)
                {
                    if (result.Value is Dictionary<string, object> auditData)
                    {
                        var riskLevel = auditData.ContainsKey("Risk Level") ? auditData["Risk Level"].ToString() : "Unknown";
                        var count = auditData.ContainsKey("Count") ? auditData["Count"].ToString() : "N/A";
                        
                        var listItem = new ListViewItem(result.Key);
                        listItem.SubItems.Add(count);
                        listItem.SubItems.Add(riskLevel);
                        summaryListView.Items.Add(listItem);
                    }
                }

                // Populate detailed results with comprehensive findings
                foreach (var result in comprehensiveResults)
                {
                    if (result.Value is Dictionary<string, object> auditData)
                    {
                        if (auditData.ContainsKey("Issues") && auditData["Issues"] is List<string> issues)
                        {
                            foreach (var issue in issues)
                            {
                                var item = new ListViewItem(result.Key);
                                item.SubItems.Add("Comprehensive Audit");
                                item.SubItems.Add(issue);
                                item.SubItems.Add("N/A");
                                item.SubItems.Add("N/A");
                                item.SubItems.Add("Active");
                                auditListView.Items.Add(item);
                            }
                        }
                    }
                }

                MessageBox.Show($"Comprehensive security audit completed. Analyzed {comprehensiveResults.Count} security categories.", 
                    "Comprehensive Audit Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error running comprehensive security audit: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowWeakPasswords(ListView auditListView)
        {
            try
            {
                auditListView.Items.Clear();
                var users = _auditService.GetUsersWithWeakPasswords();

                foreach (var user in users)
                {
                    var item = new ListViewItem(user.Username);
                    item.SubItems.Add(user.DisplayName);
                    item.SubItems.Add(string.Join("; ", user.SecurityIssues));
                    item.SubItems.Add(user.PasswordAge?.ToString() ?? "N/A");
                    item.SubItems.Add(user.LastLogon?.ToString("yyyy-MM-dd") ?? "Never");
                    item.SubItems.Add(user.IsEnabled ? "Enabled" : "Disabled");
                    auditListView.Items.Add(item);
                }

                MessageBox.Show($"Found {users.Count} users with weak passwords.", "Weak Passwords", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving weak passwords: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowPrivilegedAccounts(ListView auditListView)
        {
            try
            {
                auditListView.Items.Clear();
                var users = _auditService.GetPrivilegedAccounts();

                foreach (var user in users)
                {
                    var item = new ListViewItem(user.Username);
                    item.SubItems.Add(user.DisplayName);
                    item.SubItems.Add(string.Join("; ", user.SecurityIssues));
                    item.SubItems.Add(user.PasswordAge?.ToString() ?? "N/A");
                    item.SubItems.Add(user.LastLogon?.ToString("yyyy-MM-dd") ?? "Never");
                    item.SubItems.Add(user.IsEnabled ? "Enabled" : "Disabled");
                    auditListView.Items.Add(item);
                }

                MessageBox.Show($"Found {users.Count} privileged accounts.", "Privileged Accounts", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving privileged accounts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowInactiveAccounts(ListView auditListView)
        {
            try
            {
                auditListView.Items.Clear();
                var users = _auditService.GetInactiveAccounts();

                foreach (var user in users)
                {
                    var item = new ListViewItem(user.Username);
                    item.SubItems.Add(user.DisplayName);
                    item.SubItems.Add(string.Join("; ", user.SecurityIssues));
                    item.SubItems.Add(user.PasswordAge?.ToString() ?? "N/A");
                    item.SubItems.Add(user.LastLogon?.ToString("yyyy-MM-dd") ?? "Never");
                    item.SubItems.Add(user.IsEnabled ? "Enabled" : "Disabled");
                    auditListView.Items.Add(item);
                }

                MessageBox.Show($"Found {users.Count} inactive accounts.", "Inactive Accounts", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving inactive accounts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowLockedAccounts(ListView auditListView)
        {
            try
            {
                auditListView.Items.Clear();
                var users = _auditService.GetLockedAccounts();

                foreach (var user in users)
                {
                    var item = new ListViewItem(user.Username);
                    item.SubItems.Add(user.DisplayName);
                    item.SubItems.Add(string.Join("; ", user.SecurityIssues));
                    item.SubItems.Add(user.PasswordAge?.ToString() ?? "N/A");
                    item.SubItems.Add(user.LastLogon?.ToString("yyyy-MM-dd") ?? "Never");
                    item.SubItems.Add("Locked");
                    auditListView.Items.Add(item);
                }

                MessageBox.Show($"Found {users.Count} locked accounts.", "Locked Accounts", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving locked accounts: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportAuditReport(ListView auditListView)
        {
            try
            {
                if (auditListView.Items.Count == 0)
                {
                    MessageBox.Show("No audit data to export. Please run a security audit first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = $"AD_Security_Audit_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using var package = new OfficeOpenXml.ExcelPackage();
                    var worksheet = package.Workbook.Worksheets.Add("Security Audit");

                    // Add watermark
                    AddWatermarkToExcel(worksheet);

                    // Add headers (shifted down by 1 row to make room for watermark)
                    for (int i = 0; i < auditListView.Columns.Count; i++)
                    {
                        worksheet.Cells[2, i + 1].Value = auditListView.Columns[i].Text;
                        worksheet.Cells[2, i + 1].Style.Font.Bold = true;
                    }

                    // Add data (shifted down by 1 row to account for watermark)
                    for (int row = 0; row < auditListView.Items.Count; row++)
                    {
                        var item = auditListView.Items[row];
                        for (int col = 0; col < item.SubItems.Count; col++)
                        {
                            worksheet.Cells[row + 3, col + 1].Value = item.SubItems[col].Text;
                        }
                    }

                    // Auto-fit columns
                    worksheet.Cells.AutoFitColumns();

                    package.SaveAs(new FileInfo(saveFileDialog.FileName));
                    MessageBox.Show($"Security audit report exported successfully to: {saveFileDialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting audit report: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportAuditReportToPDF(ListView auditListView)
        {
            try
            {
                if (auditListView.Items.Count == 0)
                {
                    MessageBox.Show("No audit data to export. Please run a security audit first.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using var saveFileDialog = new SaveFileDialog
                {
                    Filter = "PDF Files (*.pdf)|*.pdf",
                    FileName = $"AD_Security_Audit_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using var fileStream = new FileStream(saveFileDialog.FileName, FileMode.Create);
                    using var document = new Document(PageSize.A4, 25, 25, 30, 30);
                    var writer = PdfWriter.GetInstance(document, fileStream);

                    document.Open();

                    // Add title
                    var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18);
                    var title = new Paragraph("Active Directory Security Audit Report", titleFont);
                    title.Alignment = Element.ALIGN_CENTER;
                    title.SpacingAfter = 20f;
                    document.Add(title);

                    // Add generation date
                    var dateFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);
                    var dateText = new Paragraph($"Generated on: {DateTime.Now:yyyy-MM-dd HH:mm:ss}", dateFont);
                    dateText.Alignment = Element.ALIGN_CENTER;
                    dateText.SpacingAfter = 20f;
                    document.Add(dateText);

                    // Add summary section
                    var summaryFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 14);
                    var summaryTitle = new Paragraph("Audit Summary", summaryFont);
                    summaryTitle.SpacingAfter = 10f;
                    document.Add(summaryTitle);

                    var summaryText = new Paragraph($"Total records analyzed: {auditListView.Items.Count}", dateFont);
                    summaryText.SpacingAfter = 15f;
                    document.Add(summaryText);

                    // Create table for audit results
                    var table = new PdfPTable(auditListView.Columns.Count);
                    table.WidthPercentage = 100;

                    // Add headers
                    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                    foreach (ColumnHeader column in auditListView.Columns)
                    {
                        var cell = new PdfPCell(new Phrase(column.Text, headerFont));
                        cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                        cell.Padding = 5f;
                        table.AddCell(cell);
                    }

                    // Add data rows
                    var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);
                    foreach (ListViewItem item in auditListView.Items)
                    {
                        foreach (ListViewItem.ListViewSubItem subItem in item.SubItems)
                        {
                            var cell = new PdfPCell(new Phrase(subItem.Text, dataFont));
                            cell.Padding = 3f;
                            table.AddCell(cell);
                        }
                    }

                    document.Add(table);

                    // Add watermark
                    AddWatermarkToPDF(document, writer);

                    // Add footer
                    document.Add(new Paragraph(" ")); // Spacing
                    var footerFont = FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8);
                    var footer = new Paragraph("This report was generated by SonicADRecon", footerFont);
                    footer.Alignment = Element.ALIGN_CENTER;
                    document.Add(footer);

                    document.Close();
                    writer.Close();

                    MessageBox.Show($"Security audit report exported successfully to: {saveFileDialog.FileName}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting audit report to PDF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
} 