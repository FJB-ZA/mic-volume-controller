using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using NAudio.CoreAudioApi;

namespace MicVolumeController
{
	/// <summary>
	/// Main form for the Microphone Volume Controller application.
	/// Allows users to select a microphone and lock its volume to a specific level.
	/// The volume is continuously enforced to prevent other applications from changing it.
	/// </summary>
	public partial class MainForm : Form
	{
		/// <summary>
		/// Enumerator for discovering audio devices on the system.
		/// </summary>
		private MMDeviceEnumerator deviceEnumerator;

		/// <summary>
		/// The currently selected microphone device.
		/// </summary>
		private MMDevice selectedDevice;

		/// <summary>
		/// Timer that continuously enforces the volume setting every 500ms.
		/// </summary>
		private System.Windows.Forms.Timer volumeTimer;

		/// <summary>
		/// System tray icon for the application.
		/// </summary>
		private NotifyIcon trayIcon;

		/// <summary>
		/// Flag indicating whether the app should minimize to tray when closed.
		/// </summary>
		private bool closeToTray = true;

		/// <summary>
		/// TrackBar control for adjusting volume with a slider.
		/// </summary>
		private TrackBar trackVolume;

		/// <summary>
		/// NumericUpDown control for precise volume input.
		/// </summary>
		private NumericUpDown numVolume;

		/// <summary>
		/// Flag to prevent infinite loops when synchronizing volume controls.
		/// </summary>
		private bool isUpdating = false;

		/// <summary>
		/// Flag indicating whether the form is still initializing.
		/// Prevents settings from being saved during initial load.
		/// </summary>
		private bool isInitializing = true;

		/// <summary>
		/// Dictionary storing volume levels for each microphone.
		/// Key: Microphone friendly name, Value: Volume level (0-100).
		/// </summary>
		private Dictionary<string, int> microphoneVolumes = new Dictionary<string, int>();

		/// <summary>
		/// Initializes a new instance of the MainForm class.
		/// Sets up the UI, loads settings, and initializes audio devices.
		/// </summary>
		public MainForm()
		{
			InitializeComponent();
			InitializeAudio();
			LoadSettings();
			SetupStartupOption();
			SetupTrayIcon();
			isInitializing = false;
		}

		/// <summary>
		/// Initializes all UI components and controls.
		/// Creates the form layout, controls, and event handlers.
		/// </summary>
		private void InitializeComponent()
		{
			this.Text = "Mic Volume Controller";
			this.Size = new System.Drawing.Size(400, 260);
			this.FormBorderStyle = FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.StartPosition = FormStartPosition.CenterScreen;

			// Set the form icon (taskbar icon)
			try
			{
				this.Icon = new System.Drawing.Icon("mic_controller_icon.ico");
			}
			catch
			{
				// If icon file not found, use default
			}

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

		/// <summary>
		/// Initializes the audio subsystem and loads available microphones.
		/// </summary>
		private void InitializeAudio()
		{
			deviceEnumerator = new MMDeviceEnumerator();
			LoadMicrophones();
		}

		/// <summary>
		/// Loads all active capture (microphone) devices into the dropdown list.
		/// </summary>
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

		/// <summary>
		/// Loads saved volume levels for all microphones from application settings.
		/// Volumes are stored as a semicolon-separated string in format "MicName1=Volume1;MicName2=Volume2".
		/// </summary>
		private void LoadMicrophoneVolumes()
		{
			// Load saved volumes from settings
			string savedVolumes = Properties.Settings.Default.MicrophoneVolumes;
			microphoneVolumes.Clear();

			if (!string.IsNullOrEmpty(savedVolumes))
			{
				// Format: "MicName1=Volume1;MicName2=Volume2"
				var pairs = savedVolumes.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
				foreach (var pair in pairs)
				{
					var parts = pair.Split('=');
					if (parts.Length == 2 && int.TryParse(parts[1], out int volume))
					{
						microphoneVolumes[parts[0]] = volume;
					}
				}
			}
		}

		/// <summary>
		/// Saves all microphone volume levels to application settings.
		/// Volumes are stored as a semicolon-separated string in format "MicName1=Volume1;MicName2=Volume2".
		/// </summary>
		private void SaveMicrophoneVolumes()
		{
			// Save volumes to settings
			// Format: "MicName1=Volume1;MicName2=Volume2"
			var volumeStrings = microphoneVolumes.Select(kvp => $"{kvp.Key}={kvp.Value}");
			Properties.Settings.Default.MicrophoneVolumes = string.Join(";", volumeStrings);
		}

		/// <summary>
		/// Gets the saved volume level for a specific microphone.
		/// </summary>
		/// <param name="micName">The friendly name of the microphone.</param>
		/// <returns>The saved volume level (0-100), or 50 if no saved value exists.</returns>
		private int GetVolumeForMicrophone(string micName)
		{
			// Return saved volume for this mic, or default to 50
			return microphoneVolumes.ContainsKey(micName) ? microphoneVolumes[micName] : 50;
		}

		/// <summary>
		/// Loads all saved application settings including microphone volumes,
		/// selected microphone, and user preferences.
		/// </summary>
		private void LoadSettings()
		{
			try
			{
				// Load microphone volumes
				LoadMicrophoneVolumes();

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

		/// <summary>
		/// Saves all current application settings including the current microphone's volume,
		/// selected microphone, and user preferences.
		/// </summary>
		private void SaveSettings()
		{
			try
			{
				// Save current microphone's volume
				if (selectedDevice != null)
				{
					microphoneVolumes[selectedDevice.FriendlyName] = (int)numVolume.Value;
				}

				SaveMicrophoneVolumes();
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

		/// <summary>
		/// Handles the microphone selection change event.
		/// Loads the saved volume for the selected microphone and starts volume monitoring.
		/// </summary>
		/// <param name="sender">The ComboBox that triggered the event.</param>
		/// <param name="e">Event arguments.</param>
		private void CmbMicrophones_SelectedIndexChanged(object sender, EventArgs e)
		{
			ComboBox cmb = (ComboBox)sender;
			if (cmb.SelectedItem != null)
			{
				var item = (MicrophoneItem)cmb.SelectedItem;
				selectedDevice = item.Device;

				// Load saved volume for this specific microphone
				int savedVolume = GetVolumeForMicrophone(item.Name);

				isUpdating = true;
				trackVolume.Value = savedVolume;
				numVolume.Value = savedVolume;
				isUpdating = false;

				// Set the volume
				SetMicrophoneVolume(savedVolume);

				// Start monitoring
				volumeTimer.Start();

				// Save settings if not initializing
				if (!isInitializing)
				{
					SaveSettings();
				}
			}
		}

		/// <summary>
		/// Handles the TrackBar value change event.
		/// Updates the numeric input and applies the volume to the microphone.
		/// </summary>
		/// <param name="sender">The TrackBar that triggered the event.</param>
		/// <param name="e">Event arguments.</param>
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

		/// <summary>
		/// Handles the NumericUpDown value change event.
		/// Updates the trackbar and applies the volume to the microphone.
		/// </summary>
		/// <param name="sender">The NumericUpDown that triggered the event.</param>
		/// <param name="e">Event arguments.</param>
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

		/// <summary>
		/// Timer tick event that continuously enforces the volume setting.
		/// This prevents other applications from changing the microphone volume.
		/// Runs every 500ms.
		/// </summary>
		/// <param name="sender">The Timer that triggered the event.</param>
		/// <param name="e">Event arguments.</param>
		private void VolumeTimer_Tick(object sender, EventArgs e)
		{
			// Continuously enforce the volume setting
			SetMicrophoneVolume((int)numVolume.Value);
		}

		/// <summary>
		/// Sets the volume of the selected microphone to the specified level.
		/// </summary>
		/// <param name="volumeLevel">The volume level to set (0-100).</param>
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

		/// <summary>
		/// Checks if the application is set to start with Windows and updates the checkbox.
		/// </summary>
		private void SetupStartupOption()
		{
			CheckBox chk = (CheckBox)this.Controls["chkStartup"];
			chk.Checked = IsInStartup();
		}

		/// <summary>
		/// Initializes the system tray icon and its context menu.
		/// </summary>
		private void SetupTrayIcon()
		{
			trayIcon = new NotifyIcon();
			trayIcon.Text = "Mic Volume Controller";

			// Use the custom icon for tray
			try
			{
				trayIcon.Icon = new System.Drawing.Icon("mic_controller_icon.ico");
			}
			catch
			{
				// If icon file not found, use default
				trayIcon.Icon = System.Drawing.SystemIcons.Application;
			}

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

		/// <summary>
		/// Handles the form closing event.
		/// If "Close to tray" is enabled, minimizes the app to the system tray instead of closing.
		/// </summary>
		/// <param name="sender">The form that triggered the event.</param>
		/// <param name="e">Event arguments that can be used to cancel the close operation.</param>
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

		/// <summary>
		/// Restores the application window from the system tray.
		/// </summary>
		private void ShowFromTray()
		{
			this.Show();
			this.WindowState = FormWindowState.Normal;
			this.BringToFront();
			trayIcon.Visible = false;
		}

		/// <summary>
		/// Exits the application completely, closing all resources.
		/// </summary>
		private void ExitApplication()
		{
			closeToTray = false;
			trayIcon.Visible = false;
			Application.Exit();
		}

		/// <summary>
		/// Handles the "Close to tray" checkbox change event.
		/// </summary>
		/// <param name="sender">The checkbox that triggered the event.</param>
		/// <param name="e">Event arguments.</param>
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

		/// <summary>
		/// Handles the "Start with Windows" checkbox change event.
		/// </summary>
		/// <param name="sender">The checkbox that triggered the event.</param>
		/// <param name="e">Event arguments.</param>
		private void ChkStartup_CheckedChanged(object sender, EventArgs e)
		{
			CheckBox chk = (CheckBox)sender;
			SetStartup(chk.Checked);
		}

		/// <summary>
		/// Checks if the application is registered to start with Windows.
		/// </summary>
		/// <returns>True if the application is set to start with Windows, false otherwise.</returns>
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

		/// <summary>
		/// Enables or disables the application to start with Windows.
		/// Modifies the Windows Registry to add or remove the startup entry.
		/// </summary>
		/// <param name="enable">True to enable startup with Windows, false to disable.</param>
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

		/// <summary>
		/// Disposes of resources used by the form.
		/// Stops the timer and disposes of audio resources.
		/// </summary>
		/// <param name="disposing">True if disposing managed resources.</param>
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

	/// <summary>
	/// Represents a microphone device in the dropdown list.
	/// Wraps an NAudio MMDevice with a user-friendly display name.
	/// </summary>
	public class MicrophoneItem
	{
		/// <summary>
		/// Gets or sets the underlying audio device.
		/// </summary>
		public MMDevice Device { get; set; }

		/// <summary>
		/// Gets or sets the friendly name of the microphone.
		/// </summary>
		public string Name { get; set; }

		/// <summary>
		/// Returns the friendly name of the microphone for display in the UI.
		/// </summary>
		/// <returns>The friendly name of the microphone.</returns>
		public override string ToString()
		{
			return Name;
		}
	}
}