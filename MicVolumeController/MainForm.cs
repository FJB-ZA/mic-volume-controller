using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;

namespace MicVolumeController
{
	public partial class MainForm : Form
	{
		private MMDeviceEnumerator deviceEnumerator;
		private MMDevice selectedDevice;
		private System.Windows.Forms.Timer volumeTimer;
		private NotifyIcon trayIcon;
		private bool closeToTray = true;
		private TrackBar trackVolume;
		private NumericUpDown numVolume;
		private bool isUpdating = false;
		private bool isInitializing = true;

		public MainForm()
		{
			InitializeComponent();
			InitializeAudio();
			LoadSettings();
			SetupStartupOption();
			SetupTrayIcon();
			isInitializing = false;
		}

		private void InitializeComponent()
		{
			this.Text = "Mic Volume Controller";
			this.Size = new System.Drawing.Size(400, 260);
			this.FormBorderStyle = FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.StartPosition = FormStartPosition.CenterScreen;

			// Microphone selection
			Label lblMic = new Label();
			lblMic.Text = "Select Microphone:";
			lblMic.Location = new System.Drawing.Point(20, 20);
			lblMic.AutoSize = true;
			this.Controls.Add(lblMic);

			ComboBox cmbMicrophones = new ComboBox();
			cmbMicrophones.Name = "cmbMicrophones";
			cmbMicrophones.Location = new System.Drawing.Point(20, 45);
			cmbMicrophones.Size = new System.Drawing.Size(340, 25);
			cmbMicrophones.DropDownStyle = ComboBoxStyle.DropDownList;
			cmbMicrophones.SelectedIndexChanged += CmbMicrophones_SelectedIndexChanged;
			this.Controls.Add(cmbMicrophones);

			// Volume slider
			Label lblVolume = new Label();
			lblVolume.Text = "Set Volume Level:";
			lblVolume.Location = new System.Drawing.Point(20, 85);
			lblVolume.AutoSize = true;
			this.Controls.Add(lblVolume);

			trackVolume = new TrackBar();
			trackVolume.Location = new System.Drawing.Point(20, 110);
			trackVolume.Size = new System.Drawing.Size(240, 45);
			trackVolume.Minimum = 0;
			trackVolume.Maximum = 100;
			trackVolume.TickFrequency = 10;
			trackVolume.Value = 50;
			trackVolume.ValueChanged += TrackVolume_ValueChanged;
			this.Controls.Add(trackVolume);

			numVolume = new NumericUpDown();
			numVolume.Location = new System.Drawing.Point(270, 115);
			numVolume.Size = new System.Drawing.Size(60, 23);
			numVolume.Minimum = 0;
			numVolume.Maximum = 100;
			numVolume.Value = 50;
			numVolume.ValueChanged += NumVolume_ValueChanged;
			this.Controls.Add(numVolume);

			Label lblPercent = new Label();
			lblPercent.Text = "%";
			lblPercent.Location = new System.Drawing.Point(335, 117);
			lblPercent.AutoSize = true;
			this.Controls.Add(lblPercent);

			// Start with Windows checkbox
			CheckBox chkStartup = new CheckBox();
			chkStartup.Name = "chkStartup";
			chkStartup.Text = "Start with Windows";
			chkStartup.Location = new System.Drawing.Point(20, 155);
			chkStartup.AutoSize = true;
			chkStartup.CheckedChanged += ChkStartup_CheckedChanged;
			this.Controls.Add(chkStartup);

			// Close to tray checkbox
			CheckBox chkCloseToTray = new CheckBox();
			chkCloseToTray.Name = "chkCloseToTray";
			chkCloseToTray.Text = "Close to system tray";
			chkCloseToTray.Location = new System.Drawing.Point(20, 180);
			chkCloseToTray.AutoSize = true;
			chkCloseToTray.Checked = true;
			chkCloseToTray.CheckedChanged += ChkCloseToTray_CheckedChanged;
			this.Controls.Add(chkCloseToTray);

			// Initialize timer for continuous volume control
			volumeTimer = new System.Windows.Forms.Timer();
			volumeTimer.Interval = 500; // Check every 500ms
			volumeTimer.Tick += VolumeTimer_Tick;

			// Handle form closing event
			this.FormClosing += MainForm_FormClosing;
		}

		private void InitializeAudio()
		{
			deviceEnumerator = new MMDeviceEnumerator();
			LoadMicrophones();
		}

		private void LoadMicrophones()
		{
			ComboBox cmb = (ComboBox)this.Controls["cmbMicrophones"];
			cmb.Items.Clear();

			var devices = deviceEnumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);

			foreach (var device in devices)
			{
				cmb.Items.Add(new MicrophoneItem { Device = device, Name = device.FriendlyName });
			}

			// Don't auto-select - let LoadSettings handle it
		}

		private void LoadSettings()
		{
			try
			{
				// Load volume setting
				int savedVolume = Properties.Settings.Default.VolumeLevel;
				if (savedVolume >= 0 && savedVolume <= 100)
				{
					trackVolume.Value = savedVolume;
					numVolume.Value = savedVolume;
				}

				// Load close to tray setting
				CheckBox chkCloseToTray = (CheckBox)this.Controls["chkCloseToTray"];
				chkCloseToTray.Checked = Properties.Settings.Default.CloseToTray;
				closeToTray = chkCloseToTray.Checked;

				// Load and select saved microphone
				string savedMicName = Properties.Settings.Default.SelectedMicrophone;
				ComboBox cmb = (ComboBox)this.Controls["cmbMicrophones"];

				if (!string.IsNullOrEmpty(savedMicName) && cmb.Items.Count > 0)
				{
					// Try to find the saved microphone
					for (int i = 0; i < cmb.Items.Count; i++)
					{
						var item = (MicrophoneItem)cmb.Items[i];
						if (item.Name == savedMicName)
						{
							cmb.SelectedIndex = i;
							return;
						}
					}
				}

				// If saved mic not found or no saved settings, select first item
				if (cmb.Items.Count > 0)
				{
					cmb.SelectedIndex = 0;
				}
			}
			catch
			{
				// If settings loading fails, use defaults
				ComboBox cmb = (ComboBox)this.Controls["cmbMicrophones"];
				if (cmb.Items.Count > 0)
				{
					cmb.SelectedIndex = 0;
				}
			}
		}

		private void SaveSettings()
		{
			try
			{
				Properties.Settings.Default.VolumeLevel = (int)numVolume.Value;
				Properties.Settings.Default.CloseToTray = closeToTray;

				ComboBox cmb = (ComboBox)this.Controls["cmbMicrophones"];
				if (cmb.SelectedItem != null)
				{
					var item = (MicrophoneItem)cmb.SelectedItem;
					Properties.Settings.Default.SelectedMicrophone = item.Name;
				}

				Properties.Settings.Default.Save();
			}
			catch
			{
				// Ignore save errors
			}
		}

		private void CmbMicrophones_SelectedIndexChanged(object sender, EventArgs e)
		{
			ComboBox cmb = (ComboBox)sender;
			if (cmb.SelectedItem != null)
			{
				var item = (MicrophoneItem)cmb.SelectedItem;
				selectedDevice = item.Device;

				// Set initial volume
				SetMicrophoneVolume((int)numVolume.Value);

				// Start monitoring
				volumeTimer.Start();

				// Save settings if not initializing
				if (!isInitializing)
				{
					SaveSettings();
				}
			}
		}

		private void TrackVolume_ValueChanged(object sender, EventArgs e)
		{
			if (isUpdating) return;

			isUpdating = true;
			numVolume.Value = trackVolume.Value;
			SetMicrophoneVolume(trackVolume.Value);
			isUpdating = false;

			// Save settings if not initializing
			if (!isInitializing)
			{
				SaveSettings();
			}
		}

		private void NumVolume_ValueChanged(object sender, EventArgs e)
		{
			if (isUpdating) return;

			isUpdating = true;
			trackVolume.Value = (int)numVolume.Value;
			SetMicrophoneVolume((int)numVolume.Value);
			isUpdating = false;

			// Save settings if not initializing
			if (!isInitializing)
			{
				SaveSettings();
			}
		}

		private void VolumeTimer_Tick(object sender, EventArgs e)
		{
			// Continuously enforce the volume setting
			SetMicrophoneVolume((int)numVolume.Value);
		}

		private void SetMicrophoneVolume(int volumeLevel)
		{
			if (selectedDevice != null)
			{
				try
				{
					float volume = volumeLevel / 100f;
					selectedDevice.AudioEndpointVolume.MasterVolumeLevelScalar = volume;
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error setting volume: " + ex.Message, "Error",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void SetupStartupOption()
		{
			CheckBox chk = (CheckBox)this.Controls["chkStartup"];
			chk.Checked = IsInStartup();
		}

		private void SetupTrayIcon()
		{
			trayIcon = new NotifyIcon();
			trayIcon.Text = "Mic Volume Controller";
			trayIcon.Icon = System.Drawing.SystemIcons.Application;
			trayIcon.Visible = false;

			// Create context menu for tray icon
			ContextMenuStrip trayMenu = new ContextMenuStrip();

			ToolStripMenuItem showItem = new ToolStripMenuItem("Show");
			showItem.Click += (s, e) => ShowFromTray();
			trayMenu.Items.Add(showItem);

			trayMenu.Items.Add(new ToolStripSeparator());

			ToolStripMenuItem exitItem = new ToolStripMenuItem("Exit");
			exitItem.Click += (s, e) => ExitApplication();
			trayMenu.Items.Add(exitItem);

			trayIcon.ContextMenuStrip = trayMenu;
			trayIcon.DoubleClick += (s, e) => ShowFromTray();
		}

		private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			CheckBox chk = (CheckBox)this.Controls["chkCloseToTray"];

			if (closeToTray && chk.Checked && e.CloseReason == CloseReason.UserClosing)
			{
				e.Cancel = true;
				this.Hide();
				trayIcon.Visible = true;
				trayIcon.ShowBalloonTip(2000, "Mic Volume Controller",
					"App minimized to tray. Double-click to restore.", ToolTipIcon.Info);
			}
		}

		private void ShowFromTray()
		{
			this.Show();
			this.WindowState = FormWindowState.Normal;
			this.BringToFront();
			trayIcon.Visible = false;
		}

		private void ExitApplication()
		{
			closeToTray = false;
			trayIcon.Visible = false;
			Application.Exit();
		}

		private void ChkCloseToTray_CheckedChanged(object sender, EventArgs e)
		{
			CheckBox chk = (CheckBox)sender;
			closeToTray = chk.Checked;

			// Save settings if not initializing
			if (!isInitializing)
			{
				SaveSettings();
			}
		}

		private void ChkStartup_CheckedChanged(object sender, EventArgs e)
		{
			CheckBox chk = (CheckBox)sender;
			SetStartup(chk.Checked);
		}

		private bool IsInStartup()
		{
			try
			{
				RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", false);
				return key?.GetValue("MicVolumeController") != null;
			}
			catch
			{
				return false;
			}
		}

		private void SetStartup(bool enable)
		{
			try
			{
				RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);

				if (enable)
				{
					key?.SetValue("MicVolumeController", Application.ExecutablePath);
				}
				else
				{
					key?.DeleteValue("MicVolumeController", false);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error setting startup: " + ex.Message, "Error",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				volumeTimer?.Stop();
				volumeTimer?.Dispose();
				trayIcon?.Dispose();
				selectedDevice?.Dispose();
				deviceEnumerator?.Dispose();
			}
			base.Dispose(disposing);
		}
	}

	public class MicrophoneItem
	{
		public MMDevice Device { get; set; }
		public string Name { get; set; }

		public override string ToString()
		{
			return Name;
		}
	}
}