using System;
using System.Collections.Generic;
using System.Linq;
using OpenMcdf;
using System.IO;


namespace OfficeModifier
{
	class Program
	{
		private static string filename = "";
		private static string filenameInDoc = "";
		private static string pathToFile = "";
		private static string fileToDelete = "";
		private static string fileToAdd = "";

		// Compound file that is under editing
		static CompoundFile cf;

		public static void PrintHelp()
		{
			Utils.HelpMenu();
		}
		static void Main(string[] args)
		{
			try
			{
				if (args.Length == 0 || args.Contains("-h"))
				{
					PrintHelp();
					return;
				}


				Dictionary<string, string> argDict = Utils.ParseArgs(args);

				if (argDict.ContainsKey("v"))
				{
					Utils.verbosity++;
				}

				if (argDict.ContainsKey("f"))
				{
					filename = argDict["f"];
				}
				else
				{
					Console.WriteLine("\n[!] Missing file (-f)\n");
					return;
				}

				if(argDict.ContainsKey("r"))
                {
					fileToDelete = argDict["r"];
                }

				if(argDict.ContainsKey("a"))
                {
					fileToAdd = argDict["a"];

					if (argDict.ContainsKey("n"))
					{
						filenameInDoc = argDict["n"];
					}
					else
					{
						Console.WriteLine("\n[!] Missing filename [-n] to name added file inside document structure.\n");
						return;
					}

                }


				// Open OLE compound file for editing
				try
				{
					cf = new CompoundFile(filename, CFSUpdateMode.Update, 0);
				}
				catch (Exception e)
				{
					Console.WriteLine("ERROR: Could not open file " + filename);
					Console.WriteLine("Please make sure this file exists and is .doc, .xls or .pub min the Office 97-2003 format.");
					Console.WriteLine();
					Console.WriteLine(e.Message);
					return;
				}

				// Modify File
				try
				{


					// Read Root Storage
					CFStorage testStorage = cf.RootStorage; // doc, xls or pub


					if (fileToDelete != "" || fileToAdd != "")
                    {
						if (fileToDelete != "")
						{
							pathToFile = fileToDelete;
						}
						else
                        {
							pathToFile = filenameInDoc;

						}
						
						List<string> pathList = null;

						// If filename contains slashes break them up in path and filename
						if (pathToFile.Contains("/"))
                        {
							pathList = pathToFile.Split('/').ToList();
							pathToFile = pathToFile.Split('/').Last();
						}

						// If filename contains backslashes break them up in path and filename
						if (pathToFile.Contains("\\"))
						{
							pathList = pathToFile.Split('\\').ToList();
							pathToFile = pathToFile.Split('\\').Last();
						}

						if (pathList != null)
                        {
							pathList.RemoveAt(pathList.Count - 1);
							foreach (var pathItem in pathList)
							{
								if (testStorage.TryGetStorage(pathItem) != null && pathItem != "") testStorage = testStorage.GetStorage(pathItem);
								else
								{
									if (fileToAdd != "" && pathItem != "")
                                    {
										testStorage.AddStorage(pathItem);
										testStorage = testStorage.GetStorage(pathItem);
									}
								}
							}

						}
						if (fileToAdd != "")
                        {
							byte[] cObj = File.ReadAllBytes(fileToAdd);
							testStorage.AddStream(pathToFile);
							testStorage.GetStream(pathToFile).SetData(cObj);
						}
						else
                        {
							testStorage.Delete(pathToFile);
						}
						
				
					}


					// Commit changes and close
					cf.Commit();
					cf.Close();
					CompoundFile.ShrinkCompoundFile(filename);
					Console.WriteLine("\n[*] OLE File Modding completed successfully!\n");
				}

				// Error handle for file not found
				catch (FileNotFoundException ex) when (ex.Message.Contains("Could not find file"))
				{
					Console.WriteLine("\n[!] Could not find path or file (-f). \n");
				}

				// Error handle when document specified and file chosen don't match
				catch (CFItemNotFound ex) when (ex.Message.Contains("was not found"))
				{
					Console.WriteLine("\n[!] Path or file not found!.\n");
				}

				// Error handle when document is not OLE/CFBF format
				catch (CFFileFormatException)
				{
					Console.WriteLine("\n[!] Incorrect filetype (-f). Must be an OLE strucutred file. OfficeModifier supports .doc, .xls, or .pub documents.\n");
				}
			}
			// Error handle for incorrect use of flags
			catch (IndexOutOfRangeException)
			{
				Console.WriteLine("\n[!] Flags (-f) and (-r) or [-a, -n] need an argument. Make sure you have provided these flags an argument.\n");
			}
		}
	}
}
