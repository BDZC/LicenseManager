#if !DEBUG
#line hidden
#endif

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Management;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Security;
using System.Xml.Linq;
using System.Globalization;
using System.IO;
using HGM.Hotbird64.Vlmcs;

// ReSharper disable once CheckNamespace

namespace HGM.Hotbird64.LicenseManager
{
	public static class Win32Error
	{
		public static string Message(uint errorCode)
		{
			var ex = new Win32Exception(unchecked((int)errorCode));
			return ex.Message;
		}
	}

	public class LicenseMachine
	{
		public struct ProductLicense
		{
			public int ServiceIndex;
			public ManagementObject License;
		}

		public struct LicenseProvider
		{
			public string FriendlyName;
			public string Version;
			public string LicenseClassName;
			public string ProductClassName;
			public string TokenClassName;
			public string ServiceName;
		}

		/// <summary>
		/// Provides information about the computer and the OS
		/// </summary>
		public class SysInfoClass
		{
			public OsInfo OsInfo;
			public CsProductInfo CsProductInfo;
			public string DiskSerialNumber;
			public IList<NicInfo> NicInfos = new List<NicInfo>();
			public ChassisInfo ChassisInfo = new ChassisInfo();
			public MotherBoardInfo MotherboardInfo = new MotherBoardInfo();
			public string BiosSerialNumber;
			public string BiosManufacturer;
		}

		public class NicInfo
		{
			public string NetConnectionId;
			public string MacAddress;
			public bool? NetEnabled;

		}

		public class ChassisInfo
		{
			public string Manufacturer;
			public string SerialNumber;
			public string SmbiosAssetTag;
		}

		public class MotherBoardInfo
		{
			public string Manufacturer;
			public string Product;
			public string SerialNumber;
			public string Version;
		}

		/// <summary>
		/// Can be used for detecting questionable configuration.
		/// E.g. license was activated before OS was installed or
		/// OS_Locale != License_Locale
		/// </summary>
		public struct OsInfo
		{
			public uint? BuildNumber;
			public string Caption;
			public string CsName;
			public DateTime? InstallDate;
			public DateTime? LocalDateTime;
			public CultureInfo Locale;
			public string SystemDrive;
			public CultureInfo OsLanguage;
			public string Version;
		}

		/// <summary>
		/// if any of this (except Version) changes,
		/// you loose your license and
		/// go to OOT or Notification
		/// </summary>
		public struct CsProductInfo
		{
			public string IdentifyingNumber;
			public string Name;
			public string Vendor;
			public string Uuid;
			public string Version;
		}

		private string computername;
		private readonly ConnectionOptions credentials = new ConnectionOptions();
		private ManagementScope scope;
		public List<ProductLicense> ProductLicenseList = new List<ProductLicense>();
		private readonly ObjectGetOptions wmiObjectOptions = new ObjectGetOptions(null, TimeSpan.MaxValue, true);
		public bool IncludeInactiveLicenses;

		private SysInfoClass sysInfo = new SysInfoClass();
		public SysInfoClass SysInfo => sysInfo;

		private string connectErrorString;
		public string ConnectErrorString => connectErrorString;

		public readonly LicenseProvider[] LicenseProvidersList =
		{
#if DEBUG
      // Non-existing service for testing purposes
      // Your application must work, even if some licensing service providers are not installed
      new LicenseProvider
	  {
		FriendlyName = "DUMMY Licensing Service",
		LicenseClassName = "DUMMYLicensingService",
		ProductClassName = "DUMMYLicensingProduct",
		TokenClassName = "DUMMYLicensingTokenActivationLicense",
		ServiceName = "DUMMYsppsvc"
	  },
#endif
      // Standard Windows licensing lervice as used in Windows NT 6.0 and up (Vista, Win7, Win8, Server 2008, Multipoint Server 2010, ... )
      new LicenseProvider
	  {
		FriendlyName = "Windows Software Protection",
		LicenseClassName = "SoftwareLicensingService",
		ProductClassName = "SoftwareLicensingProduct",
		TokenClassName = "SoftwareLicensingTokenActivationLicense",
		ServiceName = "sppsvc"
	  },

      // Office 2010 (incl. Visio and Project) uses it's own service to be able to run under Windows XP/Server 2003,
      // which do not have compatible licensing services
      //
      // Office 2013 uses its own service, when it thinks, the system provided service is not secure enough (Win7).
      // However when running in Windows 8/Server 2012 and up, it uses the system service (just to make things complicated).
      new LicenseProvider
	  {
		FriendlyName = "Office Software Protection",
		LicenseClassName = "OfficeSoftwareProtectionService",
		ProductClassName = "OfficeSoftwareProtectionProduct",
		TokenClassName = "OfficeSoftwareProtectionTokenActivationLicense",
		ServiceName = "osppsvc"
	  }
	};

		public static readonly string[] ActivationTypes =
		{
	  "All (recommended)",
	  "Active Directory",
	  "Key Management Server",
	  "Token-based"
	};

		public static string LicenseStatusText(uint licenseStatus)
		{
			switch (licenseStatus)
			{
				case 0:
					return "Unlicensed";
				case 1:
					return "Licensed";
				case 2:
					return "Initial grace";
				case 3:
					return "Additional grace";
				case 4:
					return "Non-genuine grace";
				case 5:
					return "Notification";
				case 6:
					return "Extended grace";
				default:
					return "Unknown";
			}
		}

		public void Connect()
		{
			connectErrorString = "";
			ProductLicenseList.Clear();
			sysInfo = new SysInfoClass();

			try
			{
				var temp = @"\\" + computername + @"\root\cimv2";
				scope = new ManagementScope(temp, credentials);

				try
				{
					scope.Connect();
				}
				catch (ManagementException ex)
				{
					if (ex.ErrorCode == ManagementStatus.LocalCredentials)
					{
						scope = new ManagementScope(temp);
						scope.Connect();
					}
				}
			}
			catch (COMException ex)
			{
				switch ((uint)ex.ErrorCode)
				{
					case 0x800706BA:
						throw new COMException("Computer " + ComputerName + " could not be reached. " +
											   "If you didn't misspell it's name, adjust " + ComputerName + "'s firewall settings " +
											   "to allow incoming WMI or turn the firewall temporarily off.", ex.ErrorCode);
					case 0x80070776:
						throw new COMException("Microsoft has designed DCOM in a way that it is unable to traverse " +
											   "through NAT devices. You must have a public IP Address (e.g. not 192.168.x.x " +
											   "or 10.x.x.x) to connect to a remote machine on the internet.\r\n\r\n" +
											   "You are restricted to computers in your intranet.", ex.ErrorCode);
					default:
						throw;
				}
			}
			catch (UnauthorizedAccessException ex)
			{
				var tempComputer = (ComputerName.Contains(".") ? ComputerName : ComputerName.ToUpper());
				var tempUser = UserName ?? WindowsIdentity.GetCurrent().Name;
				throw new UnauthorizedAccessException("Access denied by " + tempComputer +
													  " with the credentials you provided. Make sure\r\n\r\n" +
													  "1) You spelled your password correctly.\r\n" +
													  "2) Have administrative prvileges on " + tempComputer + "." +
													  (tempUser.Contains("\\")
														? ""
														: ("\r\n3) You are using a user name in the form DOMAIN\\Username " +
														   (ComputerName.Contains(".")
															 ? ""
															 : ("(e.g. " + tempComputer + "\\" + tempUser +
																"). This is required under some circumstances"))) + "."), ex);
			}

			var si = sysInfo;
			try
			{
				using (var osInfoObject = new ManagementObject(scope, new ManagementPath("Win32_OperatingSystem=@"), wmiObjectOptions))
				{
					try
					{
						si.OsInfo.BuildNumber = Convert.ToUInt32(osInfoObject["BuildNumber"]);
					}
					catch
					{
					}
					try
					{
						si.OsInfo.CsName = (string)osInfoObject["CSName"];
					}
					catch
					{
					}
					try
					{
						si.OsInfo.Caption = (string)osInfoObject["Caption"];
					}
					catch
					{
					}
					try
					{
						si.OsInfo.SystemDrive = (string)osInfoObject["SystemDrive"];
					}
					catch
					{
					}
					try
					{
						si.OsInfo.InstallDate = ManagementDateTimeConverter.ToDateTime((string)osInfoObject["InstallDate"]);
					}
					catch
					{
					}
					try
					{
						si.OsInfo.LocalDateTime = ManagementDateTimeConverter.ToDateTime((string)osInfoObject["LocalDateTime"]);
					}
					catch
					{
					}
					try
					{
						si.OsInfo.Version = (string)osInfoObject["Version"];
					}
					catch
					{
					}
					try
					{
						si.OsInfo.OsLanguage = new CultureInfo((int)(uint)osInfoObject["OSLanguage"]);
					}
					catch
					{
					}
					try
					{
						si.OsInfo.Locale = new CultureInfo(int.Parse((string)osInfoObject["Locale"], NumberStyles.AllowHexSpecifier));
					}
					catch
					{
					}
				}
			}
			catch
			{
			}

			try
			{
				using (var csProductClass = new ManagementClass(scope, new ManagementPath("Win32_ComputerSystemProduct"), wmiObjectOptions))
				{
					using (var csProductCollection = csProductClass.GetInstances())
					{
						var csProductObject = csProductCollection.OfType<ManagementObject>().First();

						try
						{
							si.CsProductInfo.IdentifyingNumber = (string)csProductObject["IdentifyingNumber"];
						}
						catch
						{
						}
						try
						{
							si.CsProductInfo.Name = (string)csProductObject["Name"];
						}
						catch
						{
						}
						try
						{
							si.CsProductInfo.Vendor = (string)csProductObject["Vendor"];
						}
						catch
						{
						}
						try
						{
							si.CsProductInfo.Uuid = (string)csProductObject["UUID"];
						}
						catch
						{
						}
						try
						{
							si.CsProductInfo.Version = (string)csProductObject["Version"];
						}
						catch
						{
						}
					}
				}
			}
			catch
			{
			}

			try
			{
				using (var biosClass = new ManagementClass(scope, new ManagementPath("Win32_BIOS"), wmiObjectOptions))
				{
					using (var biosCollection = biosClass.GetInstances())
					{
						var biosObject = biosCollection.OfType<ManagementObject>().First();

						try
						{
							si.BiosSerialNumber = (string)biosObject["SerialNumber"];
						}
						catch
						{
						}
						try
						{
							si.BiosManufacturer = (string)biosObject["Manufacturer"];
						}
						catch
						{
						}
					}
				}
			}
			catch
			{
			}

			try
			{
				using (var chassisClass = new ManagementClass(scope, new ManagementPath("Win32_SystemEnclosure"), wmiObjectOptions))
				{
					using (var chassisCollection = chassisClass.GetInstances())
					{
						var chassisObject = chassisCollection.OfType<ManagementObject>().First();

						try
						{
							si.ChassisInfo.SerialNumber = (string)chassisObject["SerialNumber"];
						}
						catch
						{
						}
						try
						{
							si.ChassisInfo.Manufacturer = (string)chassisObject["Manufacturer"];
						}
						catch
						{
						}
						try
						{
							si.ChassisInfo.SmbiosAssetTag = (string)chassisObject["SMBIOSAssetTag"];
						}
						catch
						{
						}
					}
				}
			}
			catch
			{
			}

			try
			{
				using (var baseBoardClass = new ManagementClass(scope, new ManagementPath("Win32_Baseboard"), wmiObjectOptions))
				{
					using (var baseBoardCollection = baseBoardClass.GetInstances())
					{
						var baseBoardObject = baseBoardCollection.OfType<ManagementObject>().First();

						try
						{
							si.MotherboardInfo.Product = (string)baseBoardObject["Product"];
						}
						catch
						{
						}
						try
						{
							si.MotherboardInfo.Manufacturer = (string)baseBoardObject["Manufacturer"];
						}
						catch
						{
						}
						try
						{
							si.MotherboardInfo.SerialNumber = (string)baseBoardObject["SerialNumber"];
						}
						catch
						{
						}
						try
						{
							si.MotherboardInfo.Version = (string)baseBoardObject["Version"];
						}
						catch
						{
						}
					}
				}
			}
			catch
			{
			}

			try
			{
				using (var nicClass = new ManagementClass(scope, new ManagementPath("Win32_NetworkAdapter"), wmiObjectOptions))
				{
					using (var nicCollection = nicClass.GetInstances())
					{
						foreach (var nicObject in nicCollection)
						{
							var ni = new NicInfo();
							try
							{
								ni.NetConnectionId = (string)nicObject["NetConnectionID"];
							}
							catch
							{
							}
							try
							{
								ni.MacAddress = (string)nicObject["MACAddress"];
							}
							catch
							{
							}
							try
							{
								ni.NetEnabled = (bool?)nicObject["NetEnabled"];
							}
							catch
							{
							}

							if (ni.NetConnectionId != null && ni.MacAddress != null)
							{
								si.NicInfos.Add(ni);
							}
						}
					}
				}
			}
			catch
			{
			}

			try
			{
				using (var disk2PartClass = new ManagementClass(scope, new ManagementPath("Win32_LogicalDiskToPartition"), wmiObjectOptions))
				{
					using (var disk2PartCollection = disk2PartClass.GetInstances())
					{
						foreach (var disk2PartObject in disk2PartCollection)
						{
							using (ManagementObject partitionObject = new ManagementObject(scope, new ManagementPath((string)disk2PartObject["Antecedent"]), wmiObjectOptions),
							  logicalDrive = new ManagementObject(scope, new ManagementPath((string)disk2PartObject["Dependent"]), wmiObjectOptions))
							{
								if ((string)logicalDrive["DeviceID"] != si.OsInfo.SystemDrive) continue;

								var x = @"Win32_DiskDrive.DeviceID='\\.\PHYSICALDRIVE" + partitionObject["DiskIndex"] + "'";
								using (var physicalDisk = new ManagementObject(scope, new ManagementPath(x), wmiObjectOptions))
								{
									si.DiskSerialNumber = ((string)physicalDisk["SerialNumber"]);
									if (si.DiskSerialNumber != null)
									{
										var serial = si.DiskSerialNumber;
										var decodedSerial = "";
										try
										{
											for (var i = 0; i < serial.Length; i += 4)
											{
												for (var j = 2; j >= 0; j -= 2)
												{
													var hexByteString = serial.Substring(i + j, 2);
													decodedSerial += (char)ushort.Parse(hexByteString, NumberStyles.AllowHexSpecifier);
												}
											}
											si.DiskSerialNumber = decodedSerial.Trim();
										}
										catch
										{
										}
									}
								}
							}
						}
					}
				}
			}
			catch
			{
			}

			for (var i = 0; i < LicenseProvidersList.Length; i++)
			{
				LicenseProvidersList[i].Version = null;
				try
				{
					var serviceClass = new ManagementClass(scope,
						new ManagementPath(LicenseProvidersList[i].LicenseClassName),
						wmiObjectOptions);
					foreach (var serviceItem in serviceClass.GetInstances())
					{
						LicenseProvidersList[i].Version = (string)serviceItem["Version"];
					}
				}
				catch (FileNotFoundException)
				{
					
				}
				catch (ManagementException ex)
				{
					switch (ex.ErrorCode)
					{
						case ManagementStatus.NotFound:
							break;
						default:
							connectErrorString += Environment.NewLine + LicenseProvidersList[i].FriendlyName + " reports: " + ex.Message;
							break;
					}
				}
				catch (Exception ex)
				{
					connectErrorString += Environment.NewLine + LicenseProvidersList[i].FriendlyName + " reports: " + ex.Message;
				}


			}

			//RefreshLicenses();
		}

		public void Connect(string localComputername, string username, string password, bool includeInactiveLicenses)
		{
			ComputerName = localComputername;
			UserName = ComputerName == "." ? null : username;
			Password = UserName == null ? null : password;
			IncludeInactiveLicenses = includeInactiveLicenses;
			Connect();
		}

		public void Connect(string localComputername, string username, SecureString securePassword, bool includeInactiveLicenses)
		{
			ComputerName = localComputername;
			UserName = ComputerName == "." ? null : username;
			SecurePassword = UserName == null ? null : securePassword;
			IncludeInactiveLicenses = includeInactiveLicenses;
			Connect();
		}

		public LicenseMachine(string computername, string username, string password, bool includeInactiveLicenses = false)
		{
			Connect(computername, username, password, includeInactiveLicenses);
		}

		public LicenseMachine(string computername, string username, SecureString securePassword, bool includeInactiveLicenses = false)
		{
			Connect(computername, username, securePassword, includeInactiveLicenses);
		}

		public LicenseMachine()
		{
			Connect(null, null, (string)null, false);
		}

		public string ComputerName
		{
			get => computername;
			private set
			{
				if (string.IsNullOrEmpty(value)) value = ".";
				computername = value;
			}
		}

		public string UserName
		{
			get => credentials.Username;
			private set
			{
				if (value == "") value = null;
				credentials.Username = value;
			}
		}

		public string Password
		{
			set => credentials.Password = value;
		}

		public SecureString SecurePassword
		{
			set => credentials.SecurePassword = value;
		}

		public void GetKmsLicenses(out bool isWindowsActivated, ICollection<KmsLicense> result)
		{
			try
			{
				using (var queryResult = GetManagementObjectCollection("SELECT LicenseStatus from SoftwareLicensingProduct WHERE ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f' and LicenseStatus = 1"))
				{
					isWindowsActivated = queryResult.OfType<ManagementObject>().Any();
				}
			}
			catch
			{
				isWindowsActivated = false;
			}

			foreach (var licenseProvider in LicenseProvidersList)
			{
				ManagementObjectCollection queryResult = null;

				try
				{
					var querystring = $"SELECT ApplicationID, Description, ID, Name, PartialProductKey, LicenseStatus from {licenseProvider.ProductClassName} WHERE Description like '%KMSCLIENT%'";

					queryResult = GetManagementObjectCollection(querystring);

					foreach (var license in queryResult)
					{
						result.Add(new KmsLicense
						{
							ApplicationID = new KmsGuid((string)license[nameof(KmsLicense.ApplicationID)]),
							ID = new KmsGuid((string)license[nameof(KmsLicense.ID)]),
							Description = (string)license[nameof(KmsLicense.Description)],
							Name = (string)license[nameof(KmsLicense.Name)],
							PartialProductKey = (string)license[nameof(KmsLicense.PartialProductKey)],
							LicenseStatus = (LicenseStatus)(uint)license[nameof(KmsLicense.LicenseStatus)],
							List = result
						});
					}
				}
				catch (ManagementException ex)
				{
					switch (ex.ErrorCode)
					{
						case ManagementStatus.InvalidClass:
						case ManagementStatus.ProviderLoadFailure:
							break;
						default:
							MessageBox.Show(ex.Message, $"Error in {licenseProvider.FriendlyName}", MessageBoxButton.OK, MessageBoxImage.Error);
							break;
					}
				}
				catch (COMException ex)
				{
					switch ((uint)ex.ErrorCode)
					{
						case 0xC004D302:
							MessageBox.Show($"{licenseProvider.FriendlyName} was rearmed. You must reboot your computer to use it.", "Reboot required", MessageBoxButton.OK, MessageBoxImage.Error);
							break;
						default:
							MessageBox.Show(ex.Message, $"Error in {licenseProvider.FriendlyName}", MessageBoxButton.OK, MessageBoxImage.Error);
							break;
					}
				}
				finally
				{
					queryResult?.Dispose();
				}
			}
		}

		private ManagementObjectCollection GetManagementObjectCollection(string querystring)
		{
			ManagementObjectCollection queryResult;
			using (var query = new ManagementObjectSearcher(scope, new ObjectQuery(querystring)))
			{
				try
				{
					queryResult = query.Get();
				}
				catch (Exception ex)
				{
					if (ex.Source == "WinMgmt") throw;
					ReEstablishConnection(ex);
					using (var query2 = new ManagementObjectSearcher(scope, new ObjectQuery(querystring)))
					{
						queryResult = query2.Get();
					}
				}
			}
			return queryResult;
		}

		public static readonly string[] RequiredProperties = new[]
		{
	  "Name", "Description","ID", "GracePeriodRemaining", "OfflineInstallationId", "LicenseStatus","LicenseStatusReason","PartialProductKey",
	  "ProductKeyChannel","GenuineStatus","EvaluationEndDate","ApplicationID","ProductKeyID","KeyManagementServicePort","KeyManagementServiceLookupDomain","KeyManagementServiceMachine",
	  "DiscoveredKeyManagementServiceMachinePort","DiscoveredKeyManagementServiceMachineName","DiscoveredKeyManagementServiceMachineIpAddress","KeyManagementServiceProductKeyID",
	  "VLRenewalInterval","VLActivationInterval","VLActivationTypeEnabled","IsKeyManagementServiceMachine","RequiredClientCount","KeyManagementServiceCurrentCount",
	  "KeyManagementServiceTotalRequests","KeyManagementServiceFailedRequests","KeyManagementServiceUnlicensedRequests","KeyManagementServiceLicensedRequests",
	  "KeyManagementServiceNonGenuineGraceRequests","KeyManagementServiceNotificationRequests","KeyManagementServiceOOBGraceRequests","KeyManagementServiceOOTGraceRequests","ExtendedGrace",
	};


		public void RefreshLicenses()
		{
			var errorMessage = "";
			ProductLicenseList.Clear();

			for (var i = 0; i < LicenseProvidersList.Length; i++)
			{
				string propertyList;

				try
				{
					using (var managementClass = new ManagementClass(scope, new ManagementPath(LicenseProvidersList[i].ProductClassName), new ObjectGetOptions(null, TimeSpan.MaxValue, true)))
					{
						var properties = managementClass.Properties.Cast<PropertyData>().Where(p => RequiredProperties.Contains(p.Name)).ToArray();
						propertyList = properties.Aggregate("", (current, managementClassProperty) => current + (managementClassProperty.Name + (managementClassProperty != properties.Last() ? ", " : "")));
					}
				}
				catch
				{
					propertyList = "*";
				}

				ManagementObjectCollection collection = null;

				try
				{
					var querystring = $"SELECT {propertyList} from "
					                  + LicenseProvidersList[i].ProductClassName
					                  + (IncludeInactiveLicenses == false
						                  ? " WHERE PartialProductKey IS NOT NULL"
						                  : "");

					collection = GetManagementObjectCollection(querystring);

					foreach (var currentlicense in collection.OfType<ManagementObject>())
					{
						ProductLicenseList.Add(new ProductLicense
						{
							ServiceIndex = i,
							License = currentlicense
						});
					}
				}
				catch (FileNotFoundException)
				{
					
				}
				catch (ManagementException ex)
				{
					switch (ex.ErrorCode)
					{
						//case ManagementStatus.NotFound:
						case ManagementStatus.InvalidClass: // Licensing Service is not installed
						case ManagementStatus.ProviderLoadFailure: // e.g. non-fully removed Office 2010 Licensing Service after De-Installation
#if DEBUG2
                            errorMessage += "\r\n\r\nThe Licensing Provider " + LicenseProvidersList[i].FriendlyName +
                                            " has encountered an error: " + ex.Message;
#endif
							break; //Not a problem, if service was not found
						default:
							errorMessage += "\r\n\r\nThe Licensing Provider " + LicenseProvidersList[i].FriendlyName +
											" has encountered an error: " + ex.Message;
							break;
					}
				}
				catch (COMException ex)
				{
					switch ((uint)ex.ErrorCode)
					{
						case 0xC004D302:
							errorMessage += "\r\n\r\nThe Licensing Provider " + LicenseProvidersList[i].FriendlyName
									+ " was rearmed. You must reboot your computer to use this service.";
							break;
						default:
							errorMessage += "\r\n\r\nThe Licensing Provider " + LicenseProvidersList[i].FriendlyName +
									" has encountered a severe error. This happens if you tamper with its store or install " +
									"license files that are not supported by that Provider. The following error has occured: " +
									"0x" + ((uint)ex.ErrorCode).ToString("X") + " " +
									Kms.StatusMessage((uint)ex.ErrorCode);
							break;
					}
				}
				finally
				{
					collection?.Dispose();
				}
			}

			if (errorMessage != "") throw new ApplicationException(errorMessage);
		}

		private void ReEstablishConnection(Exception ex)
		{
#if DEBUG
			MessageBox.Show("RECONNECTING DUE TO ACCESS DENIED BUG\r\n\r\n" + ex.Message, "FUCK!", MessageBoxButton.OK, MessageBoxImage.Error);
#endif
			var temp = @"\\" + computername + @"\root\cimv2";
			scope = new ManagementScope(temp, credentials);
			scope.Connect();
		}

		public ManagementObject GetLicenseProviderParameters(string serviceName)
		{
			var serviceClass = new ManagementClass(scope, new ManagementPath(serviceName), wmiObjectOptions);
			ManagementObjectCollection collection;
			try
			{
				collection = serviceClass.GetInstances();
			}
			catch (Exception ex)
			{
				ReEstablishConnection(ex);
				serviceClass = new ManagementClass(scope, new ManagementPath(serviceName), wmiObjectOptions);
				collection = serviceClass.GetInstances();
			}

			var result = collection.OfType<ManagementObject>().FirstOrDefault();
			if (result == null) throw new ApplicationException("Licensing service " + serviceName + " did not return any parameters.");
			return result;
		}


		public string InstallProductKey(string key)
		{
			var errorDetail = "";
			foreach (var ls in LicenseProvidersList)
			{
				if (ls.Version == null) continue;
				try
				{
					InvokeServiceMethod(ls, "InstallProductKey", new object[] { key });
					return ls.FriendlyName;
				}
				catch (ArgumentException ex)
				{
					throw new ArgumentException(ls.LicenseClassName + " reported, the format of the key \"" + key + "\" is not valid.", nameof(key), ex);
					//BUGBUG: different licensing services could accept diffent formats (at least in theory).
					//        Microsoft License Services only check, if an empty string was entered, so this is a practical solution here.
				}
				catch (ManagementException ex)
				{
					switch (ex.ErrorCode)
					{
						case ManagementStatus.InvalidClass:
						case ManagementStatus.NotFound:
							break; // Not every computer has all licensing services installed. This is not an error condition.
						default:
							throw;
					}
				}
				catch (COMException ex)
				{
					if ((uint)ex.ErrorCode == 0xC004F025)
					{
						throw new UnauthorizedAccessException("You need administrative privileges to install a product key. Run this program as administrator.", ex);
					}
					else
					{
						errorDetail += "\r\n" + ls.FriendlyName + " reports: " + ex.Message;
					}
				}

			}
			throw new ApplicationException("None of the installed licensing services did accept the product key \"" + key + "\".\r\n" + errorDetail);
		}

		private void IgnoreMethodNotImplemented(ManagementException ex)
		{
			switch (ex.ErrorCode)
			{
				case ManagementStatus.MethodNotImplemented:
					break;
				default:
					throw ex;
			}
		}

		private void SetKeyManagementServiceOverrides(string wmiServiceName, string uniqueKey, string id, string domain, string hostname, uint port)
		{
			try
			{
				try
				{
					switch (domain)
					{
						case null:
							break;
						case "":
							InvokeMethod(wmiServiceName, uniqueKey, id, "ClearKeyManagementServiceLookupDomain", null);
							break;
						default:
							InvokeMethod(wmiServiceName, uniqueKey, id, "SetKeyManagementServiceLookupDomain", new object[] { domain });
							break;
					}
				}
				catch (ManagementException ex) { IgnoreMethodNotImplemented(ex); }
				catch (ArgumentException ex)
				{
					throw new ArgumentException("The DNS domain for KMS Service Lookup is invalid.", nameof(domain), ex);
				}


				try
				{
					switch (hostname)
					{
						case null:
							break;
						case "":
							InvokeMethod(wmiServiceName, uniqueKey, id, "ClearKeyManagementServiceMachine", null);
							break;
						default:
							InvokeMethod(wmiServiceName, uniqueKey, id, "SetKeyManagementServiceMachine", new object[] { hostname });
							break;
					}
				}
				catch (ManagementException ex) { IgnoreMethodNotImplemented(ex); }
				catch (ArgumentException ex)
				{
					throw new ArgumentException("The KMS Service Hostname is invalid.", nameof(hostname), ex);
				}


				try
				{
					if (port != 0)
					{
						InvokeMethod(wmiServiceName, uniqueKey, id, "SetKeyManagementServicePort", new object[] { port });
					}
					else
					{
						InvokeMethod(wmiServiceName, uniqueKey, id, "ClearKeyManagementServicePort", null);
					}
				}
				catch (ManagementException ex) { IgnoreMethodNotImplemented(ex); }
				catch (ArgumentException ex)
				{
					throw new ArgumentException("The KMS service port must be an integer number between 1 and 65535.", nameof(port), ex);
				}
			}

			catch (UnauthorizedAccessException ex)
			{
				throw new UnauthorizedAccessException("You need administrative privileges to change the KMS service host parameters. Run this program as administrator.", ex);
			}
			catch (COMException ex)
			{
				//The f***ing license service always returns an empty error messsage
				if (ex.Source == "WinMgmt")
					throw new COMException(Kms.StatusMessage((uint)ex.ErrorCode), ex.ErrorCode);

				throw;
			}
		}

		private void GetProductLicenseId(ProductLicense productLicense, out string wmiServiceName, out string licenseId)
		{
			wmiServiceName = LicenseProvidersList[productLicense.ServiceIndex].ProductClassName;
			licenseId = (string)productLicense.License["ID"];
		}

		private void HandleServiceMethodIntParameter(LicenseProvider provider, object intObject, string methodSet, string methodClear, string errorMessage, uint defaultValue = 0)
		{
#if DEBUG
			if (methodClear == null && defaultValue == 0) throw new ApplicationException("Neither methodClear nor defaultValue was specified");
#endif
			if (intObject != null)
			{
				uint intValue = 0;
				var stringIsEmpty = (string)intObject == "";
				try
				{
					if (!stringIsEmpty) intValue = Convert.ToUInt32(intObject);
					if (intValue == 0)
					{
						if (methodClear != null)
						{
							InvokeServiceMethod(provider, methodClear, null);
						}
						else
						{
							InvokeServiceMethod(provider, methodSet, new object[] { defaultValue });
						}
					}
					else
					{
						InvokeServiceMethod(provider, methodSet, new object[] { intValue });
					}
				}
				catch (ManagementException ex) { IgnoreMethodNotImplemented(ex); }
				catch (ArgumentException ex)
				{
					throw new ArgumentException(errorMessage, ex);
				}
			}
		}

		public void SetKmsHostParameters(LicenseProvider provider, object port, object ai, object ri, bool? enableDnsPublishing = true, bool? lowPrio = false)
		{

			HandleServiceMethodIntParameter(provider, port, "SetKeyManagementServiceListeningPort",
								  "ClearKeyManagementServiceListeningPort",
								  "The KMS service port must be an integer number between 1 and 65535.");

			HandleServiceMethodIntParameter(provider, ai, "SetVLActivationInterval",
								null,
								"The activation interval must be between 15 and 43200 minutes (30 days).", 120);

			HandleServiceMethodIntParameter(provider, ri, "SetVLRenewalInterval",
								null,
								"The activation renewal interval must be between 15 an 43200 minutes (30 days).", 10080);

			if (enableDnsPublishing != null)
			{
				try
				{
					InvokeServiceMethod(provider, "DisableKeyManagementServiceDnsPublishing", new object[] { !enableDnsPublishing });
				}
				catch (ManagementException ex) { IgnoreMethodNotImplemented(ex); }
			}

			if (lowPrio != null)
			{
				try
				{
					InvokeServiceMethod(provider, "EnableKeyManagementServiceLowPriority", new object[] { lowPrio });
				}
				catch (ManagementException ex) { IgnoreMethodNotImplemented(ex); }
			}
		}

		private void SetKeyManagementOverrides(string wmiServiceName, string uniqueKey, string id, string domain, string hostname, string portstring)
		{
			uint port;
			if (string.IsNullOrEmpty(portstring))
			{
				port = 0;
			}
			else
			{
				try
				{
					port = Convert.ToUInt32(portstring);
				}
				catch (Exception ex)
				{
					throw new ApplicationException("The KMS Service port must be an integer number between 1 and 65535.", ex);
				}
			}
			SetKeyManagementServiceOverrides(wmiServiceName, uniqueKey, id, domain, hostname, port);
		}

		public void SetKeyManagementOverrides_Product(int productIndex, string domain, string hostname, object port)
		{
			SetKeyManagementOverrides_Product(ProductLicenseList[productIndex], domain, hostname, port?.ToString());
		}

		public void SetKeyManagementOverrides_Product(ProductLicense productLicense, string domain, string hostname, object port)
		{
			string productServiceName, licenseId;
			GetProductLicenseId(productLicense, out productServiceName, out licenseId);
			SetKeyManagementOverrides(productServiceName, "ID", licenseId, domain, hostname, port?.ToString());
		}

		public void SetKeyManagementOverrides_Service(int serviceIndex, string domain, string hostname, object port)
		{
			SetKeyManagementOverrides(LicenseProvidersList[serviceIndex].LicenseClassName,
						  "Version",
						  LicenseProvidersList[serviceIndex].Version,
						  domain, hostname, port?.ToString());
		}

		public void SetKeyManagementOverrides_Service(LicenseProvider licenseService, string domain, string hostname, object port)
		{
			SetKeyManagementOverrides(licenseService.LicenseClassName, "Version", licenseService.Version, domain, hostname, port?.ToString());
		}

		private object InvokeMethod(string wmiServiceName,
					  string uniqueKey,
					  string id,
					  string method,
					  object[] inParams)
		{
			try
			{
				using (var target = new ManagementObject(scope, new ManagementPath(wmiServiceName + "." + uniqueKey + "='" + id + "'"), wmiObjectOptions))
					try
					{
						return target.InvokeMethod(method, inParams);
					}
					catch (ManagementException)
					{
						throw;
					}
					catch (Exception ex)
					{
						if (ex.Source == "WinMgmt") throw;
						ReEstablishConnection(ex);
						using (var target2 = new ManagementObject(scope, new ManagementPath(wmiServiceName + "." + uniqueKey + "='" + id + "'"), wmiObjectOptions))
						{
							return target2.InvokeMethod(method, inParams);
						}
					}
			}
			catch (UnauthorizedAccessException ex)
			{
				var tempString = "You need administrative privileges to perform " + method + ". ";
				if (ComputerName == ".") tempString += "Run this program as administrator.";
				throw new UnauthorizedAccessException(tempString, ex);
			}
			catch (COMException ex)
			{
				//The f***ing license service always returns an empty error messsage
				if (ex.Source == "WinMgmt")
					throw new COMException(Kms.StatusMessage((uint)ex.ErrorCode), ex.ErrorCode);
				else
					throw;
			}
		}

		//private void InvokeProductMethod(string wmiServiceName, string licenseId, string method, object[] inParams) => InvokeMethod(wmiServiceName, "ID", licenseId, method, inParams);
		private void InvokeProductMethod(int productIndex, string method, object[] inParams) => InvokeProductMethod(ProductLicenseList[productIndex], method, inParams);

		private void InvokeProductMethod(ProductLicense productLicense, string method, object[] inParams)
		{
		    GetProductLicenseId(productLicense, out var serviceName, out var licenseId);
			InvokeMethod(serviceName, "ID", licenseId, method, inParams);
		}

		private void InvokeServiceMethod(LicenseProvider licenseProvider, string method, object[] inParams)
		{
			if (licenseProvider.Version == null)
			{
				throw new ApplicationException("The license provider " +
								 licenseProvider.FriendlyName +
								 "is not installed on the target machine.");
			}
			InvokeMethod(licenseProvider.LicenseClassName, "Version", licenseProvider.Version, method, inParams);
		}

		private void InvokeServiceMethod(int providerIndex, string method, object[] inParams) => InvokeServiceMethod(LicenseProvidersList[providerIndex], method, inParams);
		//private void Activate(string productServiceName, string licenseId) => InvokeProductMethod(productServiceName, licenseId, "Activate", null);
		public void Activate(int productIndex) => InvokeProductMethod(productIndex, "Activate", null);
		//public void Activate(ProductLicense productLicense) => InvokeProductMethod(productLicense, "Activate", null);
		//private void UninstallProductKey(string productServiceName, string licenseId) => InvokeProductMethod(productServiceName, licenseId, "UninstallProductKey", null);
		public void UninstallProductKey(int productIndex) => InvokeProductMethod(productIndex, "UninstallProductKey", null);

		public void SetVlActivationTypeEnabled(ProductLicense productLicense, uint activationType)
		{
			try
			{
				InvokeProductMethod(productLicense, "SetVLActivationTypeEnabled", new object[] { activationType });
			}
			catch { }
		}

		public void SetVlActivationTypeEnabled(int productIndex, uint activationType)
		{
			try
			{
				InvokeProductMethod(productIndex, "SetVLActivationTypeEnabled", new object[] { activationType });
			}
			catch { }
		}

		public void SetKeyManagementServiceHostCaching(LicenseProvider provider, bool enabled)
		{
			try
			{
				InvokeServiceMethod(provider,
						  "DisableKeyManagementServiceHostCaching",
						  new object[] { !enabled });
			}
			catch (ManagementException ex)
			{
				IgnoreMethodNotImplemented(ex);
			}
		}

		public void SetKeyManagementServiceHostCaching(int providerIndex, bool enabled)
		{
			SetKeyManagementServiceHostCaching(LicenseProvidersList[providerIndex], enabled);
		}

		public void InstallLicense(LicenseProvider provider, string data)
		{
			InvokeServiceMethod(provider, "InstallLicense", new object[] { data });
		}

		public void InstallLicense(int providerIndex, string data)
		{
			InvokeServiceMethod(providerIndex, "InstallLicense", new object[] { data });
		}

		public void InstallLicenseFile(LicenseProvider provider, string fileName)
		{
			var xml = XDocument.Load(fileName, LoadOptions.PreserveWhitespace);
			var xmlString = xml.ToString(SaveOptions.DisableFormatting);
			InstallLicense(provider, xmlString);
		}

		public void InstallLicenseFile(int providerIndex, string fileName)
		{
			InstallLicenseFile(LicenseProvidersList[providerIndex], fileName);
		}

		public void RefreshLicenseStatus(LicenseProvider provider)
		{
			InvokeServiceMethod(provider, "RefreshLicenseStatus", null);
		}

		public void ReArmWindows(LicenseProvider provider)
		{
			InvokeServiceMethod(provider, "ReArmWindows", null);
		}

		private ManagementObject GetServiceState(LicenseProvider provider)
		{
			ManagementObject result = null;
			try
			{
				result = new ManagementObject(scope, new ManagementPath("Win32_Service.Name='" + provider.ServiceName + "'"), wmiObjectOptions);
			}
			catch (Exception ex)
			{
				if (ex.Source == "WinMgmt") throw;
				try
				{
					ReEstablishConnection(ex);
					result = new ManagementObject(scope, new ManagementPath("Win32_Service.Name='" + provider.ServiceName + "'"), wmiObjectOptions);
				}
				catch (COMException ex2)
				{
					if (ex2.Source == "WinMgmt") throw new COMException(Kms.StatusMessage((uint)ex2.ErrorCode), ex2);
				}
			}
			return result;
		}


		public uint StopService(LicenseProvider provider, int maxTries = 3)
		{
			for (var i = 0; ; i++)
			{
				try
				{
					/*                    MessageBox.Show(service["Name"] +
									  ": State=" + service["State"] +
									  ", Status=" + service["Status"] +
									  ", Started=" + ((bool)service["Started"] == true ? "True" : "False"));*/

					var result = InvokeMethod("Win32_Service", "Name", provider.ServiceName, "StopService", null);
					switch ((uint)result)
					{
						case 0:    // Success
							break;
						case 10:   // (Has already been stopped) Never happens, because WMI class is buggy, but just in case
							break;
						case 5:    // Has already been stopped (should be 10 according to http://msdn.microsoft.com/en-us/library/windows/desktop/aa393660(v=vs.85).aspx)
							using (var service = GetServiceState(provider))
							{
								if ((bool)service["Started"]
								  || (string)service["State"] != "Stopped"
								  || (string)service["Status"] != "OK")
								{
									throw new ApplicationException(provider.ServiceName + " could not be stopped.");
								}
							}
							break;
						default:
							throw new ApplicationException(provider.ServiceName + " could not be stopped.");
					}
					return (uint)result;
				}
				catch (Exception)
				{
					if (i == maxTries) throw;
					System.Threading.Thread.Sleep(500);
				}
			}
		}

		public uint StartService(LicenseProvider provider, int maxTries = 3)
		{
			for (var i = 0; ; i++)
			{
				try
				{
					var result = InvokeMethod("Win32_Service", "Name", provider.ServiceName, "StartService", null);

					switch ((uint)result)
					{
						case 0: // Success
						case 10: // Has already been started
							break;
						default:
							throw new ApplicationException(provider.FriendlyName + " could not be started.");
					}
					return (uint)result;
				}
				catch (Exception)
				{
					if (i == maxTries) throw;
					System.Threading.Thread.Sleep(500);
				}
			}
		}
	}
}
