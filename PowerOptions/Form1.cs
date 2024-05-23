using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Configuration;

namespace PowerOptions
{
    public partial class Form1 : Form
    {
        private NotifyIcon notifyIcon1;
        private MenuItem selectPowerPlanMenuItem;
        private bool dblClickToOpen;
        private bool invertIconColor;

        public Form1()
        {
            InitializeComponent();
            LoadSettings();
            InitializeTrayIcon();
        }

        private void LoadSettings()
        {
            dblClickToOpen = Properties.Settings.Default.DblClick;
            invertIconColor = Properties.Settings.Default.InvertIconColor;
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.DblClick = dblClickToOpen;
            Properties.Settings.Default.InvertIconColor = invertIconColor;
            Properties.Settings.Default.Save();
        }

        private void InitializeTrayIcon()
        {
            notifyIcon1 = new NotifyIcon
            {
                Icon = GetTrayIcon(),
                Visible = true
            };

            notifyIcon1.MouseClick += NotifyIcon1_MouseClick;
            notifyIcon1.MouseDoubleClick += NotifyIcon1_MouseDoubleClick;

            ContextMenu trayMenu = new ContextMenu();
            MenuItem settingsMenuItem = new MenuItem("Settings");
            trayMenu.MenuItems.Add(settingsMenuItem);
            MenuItem dblClickMenuItem = new MenuItem("Double Click to open", SettingsDblClick_Click)
            {
                Checked = dblClickToOpen
            };
            settingsMenuItem.MenuItems.Add(dblClickMenuItem);
            MenuItem invertIconMenuItem = new MenuItem("Invert Icon Color", SettingsInvertIcon_Click)
            {
                Checked = invertIconColor
            };
            settingsMenuItem.MenuItems.Add(invertIconMenuItem);
            MenuItem AddNewPowerPlan = new MenuItem("Add Power Plan");
            trayMenu.MenuItems.Add(AddNewPowerPlan);
            AddNewPowerPlan.MenuItems.Add("Balanced", AddBalanced_Click);
            AddNewPowerPlan.MenuItems.Add("Power Saving", AddPowerSaving_Click);
            AddNewPowerPlan.MenuItems.Add("High Performance", AddHighPerformance_Click);
            AddNewPowerPlan.MenuItems.Add("Ultimate Performance", AddUltimatePerformance_Click);
            MenuItem SelectPowerPlan = new MenuItem("Select Power Plan");
            trayMenu.MenuItems.Add(SelectPowerPlan);
            selectPowerPlanMenuItem = SelectPowerPlan;
            GetPowerPlans(SelectPowerPlan);
            trayMenu.MenuItems.Add("Exit", Exit_Click);

            notifyIcon1.ContextMenu = trayMenu;
        }

        private Icon GetTrayIcon()
        {
            string iconPath = invertIconColor ? "iconBatteryBlack.ico" : "iconBatteryWhite.ico";
            return new Icon(Path.Combine(Application.StartupPath, "Resources", iconPath));
        }

        private void GetPowerPlans(MenuItem SelectPowerPlan)
        {
            SelectPowerPlan.MenuItems.Clear();

            string activePlanId = GetActivePowerPlan();

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = "/c powercfg /list",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                if (process != null)
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();

                    string[] lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string line in lines)
                    {
                        Match match = Regex.Match(line, @"Power Scheme GUID: ([^\s]+)\s+\((.*)\)");
                        if (match.Success)
                        {
                            string planInstanceId = match.Groups[1].Value;
                            string planName = match.Groups[2].Value;

                            MenuItem planMenuItem = new MenuItem(planName, SelectPowerPlan_Click)
                            {
                                Tag = planInstanceId
                            };

                            if (planInstanceId.Equals(activePlanId, StringComparison.OrdinalIgnoreCase))
                            {
                                planMenuItem.Checked = true;
                            }

                            SelectPowerPlan.MenuItems.Add(planMenuItem);
                        }
                    }
                }
            }
        }

        private string GetActivePowerPlan()
        {
            string activePlanId = null;

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = "/c powercfg /getactivescheme",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                if (process != null)
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();

                    string[] lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string line in lines)
                    {
                        Match match = Regex.Match(line, @"Power Scheme GUID: ([^\s]+)");
                        if (match.Success)
                        {
                            activePlanId = match.Groups[1].Value;
                            break;
                        }
                    }
                }
            }

            return activePlanId;
        }

        private void SelectPowerPlan_Click(object sender, EventArgs e)
        {
            MenuItem menuItem = (MenuItem)sender;
            string planInstanceId = menuItem.Tag.ToString();
            RunPowerCfgCommand($"powercfg /setactive {planInstanceId}");
            GetPowerPlans(menuItem.Parent as MenuItem);
        }

        private bool MessageBoxReturn(string name)
        {
            var result = MessageBox.Show($"Are you sure you want to create a “{name}” plan?", "New Power Plan", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            return result == DialogResult.Yes;
        }

        private void RunPowerCfgCommand(string command)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = $"/c {command}",
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false
            };
            Process.Start(psi);
        }

        private void CreateNewPowerPlan(int id)
        {
            switch (id)
            {
                case 0:
                    RunPowerCfgCommand("powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61");
                    break;
                case 1:
                    RunPowerCfgCommand("powercfg -duplicatescheme 381b4222-f694-41f0-9685-ff5bb260df2e");
                    break;
                case 2:
                    RunPowerCfgCommand("powercfg -duplicatescheme a1841308-3541-4fab-bc81-f71556f20b4a");
                    break;
                case 3:
                    RunPowerCfgCommand("powercfg -duplicatescheme 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c");
                    break;
                default:
                    MessageBox.Show("Error", "Can't Add new Power Plan. Make sure that the program starts with Admin rights", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }

            GetPowerPlans(selectPowerPlanMenuItem);
        }

        private void AddBalanced_Click(object sender, EventArgs e)
        {
            if (!MessageBoxReturn("Balanced")) return;
            CreateNewPowerPlan(1);
        }

        private void AddPowerSaving_Click(object sender, EventArgs e)
        {
            if (!MessageBoxReturn("Power Saving")) return;
            CreateNewPowerPlan(2);
        }

        private void AddHighPerformance_Click(object sender, EventArgs e)
        {
            if (!MessageBoxReturn("High Performance")) return;
            CreateNewPowerPlan(3);
        }

        private void AddUltimatePerformance_Click(object sender, EventArgs e)
        {
            if (!MessageBoxReturn("Ultimate Performance")) return;
            CreateNewPowerPlan(0);
        }

        private void SettingsDblClick_Click(object sender, EventArgs e)
        {
            MenuItem menuItem = (MenuItem)sender;
            menuItem.Checked = !menuItem.Checked;
            dblClickToOpen = menuItem.Checked;
            SaveSettings();
        }

        private void SettingsInvertIcon_Click(object sender, EventArgs e)
        {
            MenuItem menuItem = (MenuItem)sender;
            menuItem.Checked = !menuItem.Checked;
            invertIconColor = menuItem.Checked;
            notifyIcon1.Icon = GetTrayIcon();
            SaveSettings();
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            notifyIcon1.Dispose();
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Hide();
        }

        private void NotifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && !dblClickToOpen)
            {
                OpenPowerOptions();
            }
        }

        private void NotifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && dblClickToOpen)
            {
                OpenPowerOptions();
            }
        }

        private void OpenPowerOptions()
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "control",
                Arguments = "powercfg.cpl",
                UseShellExecute = true
            };
            Process.Start(psi);
        }
    }
}
