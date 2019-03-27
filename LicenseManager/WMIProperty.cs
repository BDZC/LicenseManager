using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Management;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using HGM.Hotbird64.Vlmcs;

// ReSharper disable once CheckNamespace
namespace HGM.Hotbird64.LicenseManager
{
  internal class WmiProperty
	{
		public string Servicename;
		public ManagementObject ManagementObject;
		public object Value { private set; get; }
		private string property;
		private readonly bool showAllFields;

		public string Property
		{
			set
			{
				property = value;

			  try
			  {
			    Value = ManagementObject[property] ?? "";
			  }
			  catch (ManagementException ex)
			  {
			    switch (ex.ErrorCode)
			    {
			      case ManagementStatus.NotFound:
			        Value = null;
			        break;
			      default:
			        throw;
			    }
			  }
			  catch
			  {
			    if (Environment.OSVersion.Version.Build > 6002) throw;
			    Value = null;
			  }
			}
			get
			{
				return property;
			}
		}

		public WmiProperty(string servicename, ManagementObject managementObject, bool showAllFields)
		{
			Servicename = servicename;
			ManagementObject = managementObject;
			this.showAllFields = showAllFields;
		}

		/*static public implicit operator CheckState(WMIProperty wmiProperty)
		{
			if ((UInt32)wmiProperty.Value == unchecked((UInt32)(-1))) return CheckState.Indeterminate;
			return (CheckState)(UInt32)wmiProperty.Value;
		}*/

		public void DisplayPropertyAsPort(TextBox textBox, string p)
		{
			Property = p;
			if (Value != null)
			{
				textBox.Text = (uint)Value == 0 ? "1688" : Value.ToString();
				Show(textBox);
			}
			else
			{
				Hide(textBox, showAllFields);
				textBox.Text = "N/A";
			}
		}

		public void DisplayAdditionalProperty(TextBox textBox, string p)
		{
			Property = p;
			if (Value != null && Value.ToString() != "")
			{
				textBox.Text += $" ({Value})";
			}
		}

		public static void Show(IEnumerable<Control> controls, TextBox textbox, bool show = true, bool showAllFields = false)
		{
			if (!show)
			{
				Hide(controls, textbox, showAllFields);
				return;
			}

			Show(textbox);

			foreach (var control in controls)
			{
				Show(control);
			}
		}

		public static void Show(Control control, TextBox textbox, bool show = true, bool showAllFields = false)
		{
			Show(new[] { control }, textbox, show, showAllFields);
		}

		public static void Show(Control control, bool show = true, bool showAllFields = false)
		{
			if (!show)
			{
				Hide(control, showAllFields);
				return;
			}

			control.Visibility = Visibility.Visible;
			control.IsEnabled = true;
		}
		public static void Hide(Control control, bool showAllFields)
		{
			if (showAllFields)
			{
				control.IsEnabled = false;
				control.Visibility = Visibility.Visible;
			}
			else
			{
				control.IsEnabled = true;
				control.Visibility = Visibility.Collapsed;
			}
		}

		public static void Hide(IEnumerable<Control> controls, TextBox textbox, bool showAllFields)
		{
			Hide(textbox, showAllFields);

			foreach (var control in controls)
			{
					Hide(control, showAllFields);
			}
		}

		public static void Hide(Control control, TextBox textbox, bool showAllFields)
		{
			Hide(new[] {control}, textbox, showAllFields);
		}

		private void Hide(IEnumerable<Control> controls, TextBox textbox)
		{
			Hide(controls, textbox, showAllFields);
		}

		private void Hide(Control control, TextBox textbox)
		{
			Hide(control, textbox, showAllFields);
		}

		[SuppressMessage("ReSharper", "CompareOfFloatsByEqualityOperator")]
		public void DisplayPropertyAsPeriodRemaining(IEnumerable<Control> controls, TextBox textbox, string p)
		{
			Property = p;

			if (Value == null)
			{
				Hide(controls, textbox);
				textbox.Text = "(unsupported in " + Servicename + ")";
				return;
			}

			Show(controls, textbox);

			var minutesRemaining = (double)(uint)Value;
			var tempDate = DateTime.Now.AddMinutes((uint)Value);
			textbox.Text = (minutesRemaining == 0.0)
					? ("forever (unless you install a new key or tamper with the license tokens)")
					: (Math.Round(minutesRemaining / 24.0 / 60.0).ToString(CultureInfo.CurrentCulture)) + " days, until " +
					tempDate.ToLongDateString() + " " + tempDate.ToShortTimeString();
		}

		public void DisplayPropertyAsPhoneId(IEnumerable<Control> controls, TextBox textbox, string p)
		{
			Property = p;

			if (Value == null || (string)Value == "")
			{
				Hide(controls, textbox);
				textbox.Text = "N/A";
				return;
			}

			Show(controls, textbox);
			try
			{
				var result = "";

				if (((string)Value).Length % 9 != 0) throw new Exception();

				for (var i = 0; i < 9; i++)
				{
					result += ((string)Value).Substring(i * ((string)Value).Length / 9, ((string)Value).Length / 9) + (i == 8 ? "" : " ");
				}

				textbox.Text = result;
			}
			catch
			{
				textbox.Text = Value.ToString();
			}

		}

		public void DisplayPropertyAsLicenseStatus(IEnumerable<Control> controls, TextBox textBox)
		{
			Property = "LicenseStatus";
			if (!(Value is uint))
			{
				Hide(controls, textBox);
				textBox.Text = "N/A";
			}
			try
			{
				var licenseStatus = (uint)Value;
				var licenseStatusString = LicenseMachine.LicenseStatusText(licenseStatus);

				switch (licenseStatus)
				{
					case 0:
					case 5:
						textBox.Background = Brushes.OrangeRed;
						break;
					case 1:
						textBox.Background = Brushes.LightGreen;
						break;
					default:
						textBox.Background = Brushes.Yellow;
						break;
				}

				Property = "LicenseStatusReason";

				if (Value != null)
				{
					licenseStatusString += ": " + Kms.StatusMessage((uint)Value);
				}

				textBox.Text = licenseStatusString;
				Show(controls, textBox);
			}
			catch
			{
				Hide(controls, textBox);
				textBox.Text = "N/A";
				textBox.Background = App.DefaultTextBoxBackground;
			}
		}

		public void DisplayProperty(IEnumerable<Control> controls, TextBox textbox, string p, bool reportUnsupported = true)
		{
			Property = p;
			if (Value == null)
			{
				Hide(controls, textbox);
				textbox.Text = reportUnsupported ? "(unsupported in " + Servicename + ")" : "";
			}
			else if (Value is uint && (uint)Value == uint.MaxValue)
			{
					textbox.Text = "N/A";
					Hide(controls, textbox);
			}
			else if (Value is string && (string)Value == "" && textbox.IsReadOnly)
			{
				Hide(controls, textbox);
				textbox.Text = "N/A";
			}
			else
			{
				textbox.Text = Value.ToString();
				Show(controls, textbox);
			}
		}

		public void DisplayProperty(Control control, TextBox textbox, string p, bool reportUnsupported = true)
		{
			DisplayProperty(new[] { control }, textbox, p, reportUnsupported);
		}

		public void DisplayPropertyAsGuid(Control control, TextBox textBox, string p, bool developerMode, bool reportUnsupported = true)
		{
			DisplayProperty(control, textBox, p, reportUnsupported);

			if (!developerMode) return;

			byte[] guidBytes;

			try
			{
				guidBytes = new Guid((string)Value).ToByteArray();
			}
			catch
			{
				return;
			}

			var data1 = (uint)(guidBytes[3] << 24 |
				guidBytes[2] << 16 |
				guidBytes[1] << 8 |
				guidBytes[0]);

			var data2 = (ushort)(guidBytes[5] << 8 |
				guidBytes[4]);

			var data3 = (ushort)(guidBytes[7] << 8 |
				guidBytes[6]);

			var byteList = "";
			string cGuid = $" / {{ 0x{data1:x8}, 0x{data2:x4}, 0x{data3:x4}, {{ ";

			for (var i = 8; i < 16; i++)
			{
				byteList += $"0x{guidBytes[i]:x2}{(i == 15 ? "" : ",")} ";
			}

			cGuid += byteList + "} } / ";
      cGuid += $"new Guid( 0x{data1:x8}, 0x{data2:x4}, 0x{data3:x4}, {byteList} )";
      textBox.Text += cGuid;
		}

		public void DisplayPropertyAsDate(IEnumerable<Control> controls, TextBox textBox, string p)
		{
			Property = p;
			try
			{
				var tempDate = ManagementDateTimeConverter.ToDateTime((string)Value);
				if (tempDate.Year == 1601)
				{
					Hide(controls, textBox);
					textBox.Text = "Never";
				}
				else
				{
					textBox.Text = tempDate.ToLongDateString() + " " + tempDate.ToLongTimeString();
					Show(controls, textBox);
				}

			}
			catch
			{
				Hide(controls, textBox);
				textBox.Text = "N/A";
			}
		}

		public void DisplayPid
		(
			Control pidControl, TextBox pidBox,
			Control osControl, TextBox osBox,
			Control dateControl, TextBox dateBox,
			string p
		)
		{
			Property = p;

			if (Value == null || (string)Value == "")
			{
				Hide(pidControl, pidBox);
				Hide(osControl, osBox);
				Hide(dateControl, dateBox);
				pidBox.Text = osBox.Text = dateBox.Text = "N/A";
				return;
			}

			var pid = new EPid(Value);
			pidBox.Text = pid.Id;
			Show(pidControl, pidBox);

			try
			{
				osBox.Text = pid.LongOsName;
				Show(osControl, osBox);
			}
			catch
			{
				Hide(osControl, osBox);
				osBox.Text = "Unknown OS";
			}

			try
			{
				dateBox.Text = pid.LongDateString;
				Show(dateControl, dateBox);
			}
			catch
			{
				Hide(dateControl, dateBox);
				dateBox.Text = "Unknown Date";
			}
		}
	
		public void SetCheckBox(CheckBox checkBox, string p)
		{
			Property = p;
			checkBox.IsChecked = (Value == null ? null : (bool?)((uint)Value == 0));
		}
	}
}
