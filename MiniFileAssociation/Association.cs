using Microsoft.Win32;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace MiniFileAssociation
{
	public class Association
	{
		private const string FileExtsRegistryPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\";
		private const string BadExtensionMessage = "The extension has to start with a '.' character.";

		#region Public methods

		/// <summary>
		/// Determines whether the extension is associated with the given .exe file.
		/// </summary>
		/// <exception cref="ArgumentException" />
		public static bool IsAssociatedWith(string extension, string exePath)
		{
			if (!extension.StartsWith("."))
				throw new ArgumentException(BadExtensionMessage);

			string openMethodKeyName = GetOpenMethodKeyName(extension);

			if (string.IsNullOrEmpty(openMethodKeyName))
				return false;

			string openCommandValue = GetOpenCommandValue(openMethodKeyName);

			if (string.IsNullOrEmpty(openCommandValue))
				return false;

			if (!openCommandValue.Equals($"\"{exePath}\" \"%1\"", StringComparison.OrdinalIgnoreCase))
				return false;

			string progId = GetProgId(extension);

			if (string.IsNullOrEmpty(progId))
				return true; // No UserChoice registry was found

			return progId == openMethodKeyName;
		}

		/// <returns>
		/// The path of the .exe file associated with the given extension
		/// or <c>null</c> if no file association has been set.
		/// </returns>
		/// <exception cref="ArgumentException" />
		public static string GetAssociatedExePath(string extension)
		{
			if (!extension.StartsWith("."))
				throw new ArgumentException(BadExtensionMessage);

			string openMethodKeyName = GetOpenMethodKeyName(extension);

			if (string.IsNullOrEmpty(openMethodKeyName))
				return null;

			string openCommandValue = GetOpenCommandValue(openMethodKeyName);
			var match = Regex.Match(openCommandValue, "^\"(.*)\"");

			return match.Success ? match.Groups[1].Value : null;
		}

		/// <summary>
		/// Sets the file association for the given extension to the given .exe file.
		/// <para>WARNING: Method requires admin privileges.</para>
		/// </summary>
		/// <param name="fileDescription">The description of the file type displayed next to the file in the Explorer.</param>
		/// <param name="iconPath">
		/// The path of the custom icon which will be given to all associated files.
		/// <para>Leave this at <c>null</c> if you want the files to inherit the program's icon.</para>
		/// </param>
		/// <exception cref="ArgumentException" />
		public static void SetAssociation(string extension, string exePath, string fileDescription, string iconPath = null)
		{
			if (!extension.StartsWith("."))
				throw new ArgumentException(BadExtensionMessage);

			if (string.IsNullOrEmpty(iconPath))
				iconPath = exePath;

			string keyName = Path.GetFileNameWithoutExtension(exePath).Replace(" ", string.Empty);

			using (RegistryKey extensionKey = Registry.ClassesRoot.CreateSubKey(extension))
				extensionKey.SetValue("", keyName);

			using (RegistryKey openMethodKey = Registry.ClassesRoot.CreateSubKey(keyName))
			{
				openMethodKey.SetValue("", fileDescription);

				using (RegistryKey defaultIconKey = openMethodKey.CreateSubKey("DefaultIcon"))
					defaultIconKey.SetValue("", $"\"{iconPath}\", 0");

				using (RegistryKey shellKey = openMethodKey.CreateSubKey("Shell"))
				using (RegistryKey openKey = shellKey.CreateSubKey("open"))
				using (RegistryKey commandKey = openKey.CreateSubKey("command"))
					commandKey.SetValue("", $"\"{exePath}\" \"%1\"");
			}

			// Delete the UserChoice key
			using (RegistryKey extensionKey = Registry.CurrentUser.OpenSubKey(FileExtsRegistryPath + extension, true))
				extensionKey?.DeleteSubKey("UserChoice", false);

			// Tell the explorer that the file association has been changed
			NativeMethods.SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero);
		}

		/// <summary>
		/// Removes all file associations for the given extension.
		/// <para>WARNING: Method requires admin privileges.</para>
		/// </summary>
		/// <exception cref="ArgumentException" />
		public static void ClearAssociations(string extension)
		{
			if (!extension.StartsWith("."))
				throw new ArgumentException(BadExtensionMessage);

			Registry.ClassesRoot.DeleteSubKey(extension, false);

			using (RegistryKey fileExtsKey = Registry.CurrentUser.OpenSubKey(FileExtsRegistryPath, true))
				fileExtsKey?.DeleteSubKeyTree(extension, false);

			// Tell the explorer that the file association has been changed
			NativeMethods.SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero);
		}

		#endregion Public methods

		#region Private methods

		private static string GetOpenMethodKeyName(string extension)
		{
			using (RegistryKey extensionKey = Registry.ClassesRoot.OpenSubKey(extension))
				return extensionKey?.GetValue("")?.ToString();
		}

		private static string GetOpenCommandValue(string openMethodkeyName)
		{
			using (RegistryKey openMethodKey = Registry.ClassesRoot.OpenSubKey(openMethodkeyName))
			using (RegistryKey shellKey = openMethodKey?.OpenSubKey("Shell"))
			using (RegistryKey openKey = shellKey?.OpenSubKey("open"))
			using (RegistryKey commandKey = openKey?.OpenSubKey("command"))
				return commandKey?.GetValue("")?.ToString();
		}

		private static string GetProgId(string extension)
		{
			using (RegistryKey extensionKey = Registry.CurrentUser.OpenSubKey(FileExtsRegistryPath + extension))
			using (RegistryKey userChoiceKey = extensionKey?.OpenSubKey("UserChoice"))
				return userChoiceKey?.GetValue("ProgId")?.ToString();
		}

		#endregion Private methods
	}
}
