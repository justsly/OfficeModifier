using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Kavod.Vba.Compression;

namespace OfficeModifier
{
    class Utils
    {

		// Verbosity level for debug messages
		public static int verbosity = 0;
		public static Dictionary<string, string> ParseArgs(string[] args)
		{
			Dictionary<string, string> ret = new Dictionary<string, string>();
			if (args.Length > 0)
			{
				for (int i = 0; i < args.Length; i += 2)
				{
					if (args[i].Substring(1).ToLower() == "l")
					{
						ret.Add(args[i].Substring(1).ToLower(), "true");
					}
					else
					{
						ret.Add(args[i].Substring(1).ToLower(), args[i + 1]);
					}
				}
			}
			return ret;
		}
		public static void HelpMenu()
		{
			Console.WriteLine("\n _____  __  __ _           ___  ___          _ _  __ _           ");
			Console.WriteLine("|  _  |/ _|/ _(_)          |  \\/  |         | (_)/ _(_)          ");
			Console.WriteLine("| | | | |_| |_ _  ___ ___  | .  . | ___   __| |_| |_ _  ___ _ __ ");
			Console.WriteLine("| | | |  _|  _| |/ __/ _ \\ | |\\/| |/ _ \\ / _` | |  _| |/ _ \\ '__|");
			Console.WriteLine("\\ \\_/ / | | | | | (_ | __/ | |  | | (_) | (_  | | | | | |__ / |");
			Console.WriteLine(" \\___/|_| |_| |_|\\___\\___| \\_|  |_/\\___/ \\__,_|_|_| |_|\\___|_|   ");
			Console.WriteLine("\n\n Author: Björn Brixner\n\n");
			Console.WriteLine(" DESCRIPTION:");
			Console.WriteLine("\n\tOfficeModifier is a C# tool to easily add or remove files from legacy Office documents. ");
			Console.WriteLine(" USAGE:");
			Console.WriteLine("\n\t-f : Filename of Office document");
			Console.WriteLine("\n\t-n : Filename to name added file inside document structure");
			Console.WriteLine("\n\t-a : file to add into document structure");
			Console.WriteLine("\n\t-r : file which should be removed from document");
			Console.WriteLine("\n\t-v : Increase debug message verbosity");
			Console.WriteLine("\n\t-h : Show help menu.\n\n");
			Console.WriteLine(" EXAMPLES:");
			Console.WriteLine("\n\t .\\OfficeModifier.exe -f .\\word.doc -a example.txt -n Macros/example.txt");
			Console.WriteLine("\n\t .\\OfficeModifier.exe -f .\\excel.xls -r _VBA_PROJECT_CUR/Project");
		}
		public static List<ModuleInformation> ParseModulesFromDirStream(byte[] dirStream)
		{
			// 2.3.4.2 dir Stream: Version Independent Project Information
			// https://msdn.microsoft.com/en-us/library/dd906362(v=office.12).aspx
			// Dir stream is ALWAYS in little endian

			List<ModuleInformation> modules = new List<ModuleInformation>();

			int offset = 0;
			UInt16 tag;
			UInt32 wLength;
			ModuleInformation currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };

			while (offset < dirStream.Length)
			{
				tag = GetWord(dirStream, offset);
				wLength = GetDoubleWord(dirStream, offset + 2);

				// taken from Pcodedmp
				if (tag == 9)
					wLength = 6;
				else if (tag == 3)
					wLength = 2;


				switch (tag)
				{
					// MODULESTREAMNAME Record
					case 26:
						currentModule.moduleName = System.Text.Encoding.UTF8.GetString(dirStream, (int)offset + 6, (int)wLength);
						break;

					// MODULEOFFSET Record
					case 49:
						currentModule.textOffset = GetDoubleWord(dirStream, offset + 6);
						modules.Add(currentModule);
						currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };
						break;
				}

				offset += 6;
				offset += (int)wLength;
			}

			return modules;
		}

		public static string HexDump(byte[] bytes, int bytesPerLine = 16)
		{
			if (bytes == null) return "<null>";
			int bytesLength = bytes.Length;

			char[] HexChars = "0123456789ABCDEF".ToCharArray();

			int firstHexColumn =
				8                   // 8 characters for the address
				+ 3;                  // 3 spaces

			int firstCharColumn = firstHexColumn
				+ bytesPerLine * 3       // - 2 digit for the hexadecimal value and 1 space
				+ (bytesPerLine - 1) / 8 // - 1 extra space every 8 characters from the 9th
				+ 2;                  // 2 spaces 

			int lineLength = firstCharColumn
				+ bytesPerLine           // - characters to show the ascii value
				+ Environment.NewLine.Length; // Carriage return and line feed (should normally be 2)

			char[] line = (new String(' ', lineLength - Environment.NewLine.Length) + Environment.NewLine).ToCharArray();
			int expectedLines = (bytesLength + bytesPerLine - 1) / bytesPerLine;
			StringBuilder result = new StringBuilder(expectedLines * lineLength);

			for (int i = 0; i < bytesLength; i += bytesPerLine)
			{
				line[0] = HexChars[(i >> 28) & 0xF];
				line[1] = HexChars[(i >> 24) & 0xF];
				line[2] = HexChars[(i >> 20) & 0xF];
				line[3] = HexChars[(i >> 16) & 0xF];
				line[4] = HexChars[(i >> 12) & 0xF];
				line[5] = HexChars[(i >> 8) & 0xF];
				line[6] = HexChars[(i >> 4) & 0xF];
				line[7] = HexChars[(i >> 0) & 0xF];

				int hexColumn = firstHexColumn;
				int charColumn = firstCharColumn;

				for (int j = 0; j < bytesPerLine; j++)
				{
					if (j > 0 && (j & 7) == 0) hexColumn++;
					if (i + j >= bytesLength)
					{
						line[hexColumn] = ' ';
						line[hexColumn + 1] = ' ';
						line[charColumn] = ' ';
					}
					else
					{
						byte b = bytes[i + j];
						line[hexColumn] = HexChars[(b >> 4) & 0xF];
						line[hexColumn + 1] = HexChars[b & 0xF];
						line[charColumn] = (b < 32 ? '·' : (char)b);
					}
					hexColumn += 3;
					charColumn++;
				}
				result.Append(line);
			}
			return result.ToString();
		}

		public static void DebugLog(object args)
		{
			if (verbosity > 0)
			{
				Console.WriteLine();
				Console.WriteLine("########## DEBUG OUTPUT: ##########");
				Console.WriteLine(args);
				Console.WriteLine("###################################");
				Console.WriteLine();
			}
		}

		public class ModuleInformation
		{
			// Name of VBA module stream
			public string moduleName;

			// Offset of VBA CompressedSourceCode in VBA module stream
			public UInt32 textOffset;
		}

		public static UInt16 GetWord(byte[] buffer, int offset)
		{
			var rawBytes = new byte[2];
			Array.Copy(buffer, offset, rawBytes, 0, 2);
			return BitConverter.ToUInt16(rawBytes, 0);
		}

		public static UInt32 GetDoubleWord(byte[] buffer, int offset)
		{
			var rawBytes = new byte[4];
			Array.Copy(buffer, offset, rawBytes, 0, 4);
			return BitConverter.ToUInt32(rawBytes, 0);
		}
		public static byte[] Compress(byte[] data)
		{
			var buffer = new DecompressedBuffer(data);
			var container = new CompressedContainer(buffer);
			return container.SerializeData();
		}
		public static byte[] Decompress(byte[] data)
		{
			var container = new CompressedContainer(data);
			var buffer = new DecompressedBuffer(container);
			return buffer.Data;
		}

		public static string ByteArrayToString(byte[] ba)
		{
			return Encoding.Default.GetString(ba);
		}

		public static string getOutFilename(String filename)
		{
			string fn = Path.GetFileNameWithoutExtension(filename);
			string ext = Path.GetExtension(filename);
			string path = Path.GetDirectoryName(filename);
			return Path.Combine(path, fn + "_MOD" + ext);
		}

		public static byte[] HexToByte(string hex)
		{
			hex = hex.Replace("-", "");
			byte[] raw = new byte[hex.Length / 2];
			for (int i = 0; i < raw.Length; i++)
			{
				raw[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
			}
			return raw;
		}
		
	}
}
