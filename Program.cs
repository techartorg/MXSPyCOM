using System;
using System.Diagnostics;
using System.Windows.Forms;


namespace MXSPyCOM
{
	class Program
	{
		/// A modern version of MXSCOM, to allow for editing & execution of 3ds Max MaxScript and Python files from external code editors.
		/// In 2005 Simon Feltman released the first MXSCOM, a small Visual Basic 6 application that took commands and sent them to 
		/// Autodesk's 3ds Max's internal COM server. This allowed users to choose their own external code editor for editing MaxScript 
		/// and to be able to have their MaxScript code execute in 3ds Max by way of having the code editor utilize MXSCOM to send the file 
		/// into 3ds Max and have it executed. Modern versions of Windows can not use Simon Feltman's old MXSCOM.exe program due to it being ActiveX based.
		/// 
		/// MXSPyCOM is a C# based replacement for MXSCOM. It offers the same functionality as MXSCOM but can run on modern versions of Windows. 
		/// It also supports editing of Python files and having them execute in versions of 3ds Max, starting with 3ds Max 2015, that support Python scripts.
		/// 
		/// **Arguments:**
		/// 
		/// None
		/// 
		/// **Keyword Arguments:**
		///
		/// None
		/// 
		/// **TODO**
		/// 
		/// :Add the ability to start 3ds Max if it is not found as a running process.
		/// 
		/// **Author:**
		/// 
		/// Jeff Hanna, jeff.b.hanna@gmail.com, July 9, 2016 9:00:00 AM

		const string USAGE_INFO = "\nType \"MXSPyCOM\" for usage info.";


		static void execute_max_commands(string[] args, string filepath)
		{
			/// Parses the command line arguments and calls the corresponding command on Max's COM server with the provied filepath.
			/// 
			/// **Arguments:**
			/// 
			/// :``args``: `string[]`: The command line arguments sent to MXSPyCOM
			/// :``filepath: `string`: A full absolute filepath to a MaxScript (.ms) or Python (.py) file.
			/// 
			/// **Keyword Arguments:**
			///
			/// None
			/// 
			/// **Returns:**
			/// 
			/// None
			/// 
			/// **TODO:**
			/// 
			/// Should support for Max's COM server execute() function be added? It doesn't act on files, but on strings of MaxScript commands.
			/// It doesn't seem necessary for this tool, honestly.
			/// 
			/// **Author:**
			/// 
			/// Jeff Hanna, jeff.b.hanna@gmail.com, July 11, 2016 9:12:00 PM

			bool max_running = is_process_running("3dsmax");
			if (max_running)
			{
				if (args.Length == 1)
				{
					string msg = String.Format("No options provided.", USAGE_INFO);
					show_message(msg);
				}
				else
				{
					var com_type = Type.GetTypeFromProgID("Max.Application");
					dynamic com_obj = Activator.CreateInstance(com_type);
					string ext = System.IO.Path.GetExtension(filepath).ToLower();

					foreach (string arg in args)
					{
						switch (arg.ToLower())
						{
							case "-f":
								
								if (ext == ".py")
								{
									filepath = make_python_wrapper(filepath);
								}

								try
								{
									com_obj.filein(filepath);
								}
								catch (System.Runtime.InteropServices.COMException) { }
								break;

							case "-e":
								try
								{									
									com_obj.edit(filepath);
								}
								catch (System.Runtime.InteropServices.COMException) { }
								break;

							case "-c":
								if (ext == ".ms")
								{
									try
									{
										com_obj.encryptscript(filepath);
									}
									catch (System.Runtime.InteropServices.COMException) { }
								}
								else
								{
									string msg = String.Format("Only MaxScript files can be encrypted. {0}", USAGE_INFO);
									show_message(msg);
								}
								break;

							default:
								return;
						}
					}
				}
			}
			else
			{
				string msg = "3ds Max (3dsmax.exe) is not currently running.";
				show_message(msg);
			}

			return;
		}


		static bool is_process_running(string process_name)
		{
			/// Determines if a named process is currently running.
			/// 
			/// **Arguments:**
			/// 
			/// :``process_name``: `string` The name of the process (minus the .exe extension) to check
			/// 
			/// **Keyword Arguments:**
			///
			/// None
			/// 
			/// **Returns:**
			/// 
			/// :`bool`
			/// 
			/// **Author:**
			/// 
			/// Jeff Hanna, jeff.b.hanna@gmail.com, July 9, 2016 9:00:00 AM 

			Process[] pname = Process.GetProcessesByName(process_name);
			if (pname.Length > 0)
			{
				return true;
			}

			return false;
		}


		static string make_python_wrapper(string python_filepath)
		{
			/// It is not possible to directly execute Python files in 3ds Max via calling filein() on the COM server.
			/// Luckily MaxScript supports Python.ExecuteFile(filepath). This function takes the provided Python file
			/// and wraps it a Python.ExecuteFile() command within a MaxScript file that is saved to the user's %TEMP% folder.
			/// That temporary MaxScript file is what is sent to 3ds Max's COM server.
			/// 
			/// **Arguments:**
			/// 
			/// :``python_filepath``: `string` A full absolute filepath to a Python (.py) file.
			/// 
			/// **Keyword Arguments:**
			///
			/// None
			/// 
			/// **Returns:**
			/// 
			/// :``wrapper_filepath``: `string`
			/// 
			/// **Author:**
			/// 
			/// Jeff Hanna, jeff.b.hanna@gmail.com, July 9, 2016 9:00:00 AM

			string cmd = String.Format("python.ExecuteFile(@\"{0}\")", python_filepath);
			string wrapper_filepath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "maxscript_python_wrapper.ms");
			System.IO.File.WriteAllText(wrapper_filepath, cmd);

			return wrapper_filepath;
		}


		static string mxs_filein_catch_errors_cmd(string filepath)
		{
			string file_line = String.Format(" Error while running: {0}", filepath);
			string mxs_exception_array = "(filterString (getCurrentException()) \"\n\")";
			string trycmd = String.Format("filein(@\"{0}\")", filepath);
			string catchcmd = String.Format("(for i in #(\"{0}\") + {1} do print i)", file_line, mxs_exception_array);
			string trycatchcmd = String.Format("try({0}) catch({1})", trycmd, catchcmd);
			return trycatchcmd;
		}


		static void show_message(string message, bool info = false)
		{
			/// Displays an error or informational dialog if the execution of MXSPyCOM encounters a problem. 
			/// Also displays a standard help dialog if MXSPyCOM is called with no arguments or with the /? or /help arguments.
			/// 
			/// **Arguments:**
			/// 
			/// :``message``: `string` The message to display.
			/// 
			/// **Keyword Arguments:**
			///
			/// :``information``: `bool` Determines if the dialog should display with an informational or warning icon.
			/// 
			/// **Returns:**
			/// 
			/// None
			/// 
			/// **Author:**
			/// 
			/// Jeff Hanna, jeff.b.hanna@gmail.com, July 9, 2016 9:00:00 AM

			MessageBoxIcon icon = info == true ? MessageBoxIcon.Information : MessageBoxIcon.Error;

			if (message.ToLower() == "help")
			{
				message = @"Used to execute MaxScript and Python scripts in 3ds Max.

Usage:
MXSPyCOM.exe [options] <filename>

Options:
-f	- filein (execute) the script in 3ds Max.
-e	- Edit the script in 3ds Max's internal script editor.
-c	- Encrypt the script. Only works with MaxScript files.

Commands:
<filename>	- Full path to the script file to execute.";

				show_message(message, info:true);
			}

			MessageBox.Show(message, "MXSPyCOM", MessageBoxButtons.OK, icon);

			Environment.Exit(0);
		}


		static string get_script_filepath(string[] args)
		{
			/// Extracts the script filepath from the command line arguments.
			/// If no filepath is provied in args the user is notifed and the program exits.
			/// If the filepath provided contains a Python script that script is wrapped in
			/// a MaxScript wrapper, because 3ds Max's COM interface will not correctly parse
			/// a Python file, and the filepath to that temporary wrapper script is returned.
			/// 
			/// **Arguments:**
			/// 
			/// :``args``: `string[]` The arguments provided to MXSPyCOM
			/// 
			/// **Keyword Arguments:**
			///
			/// None
			/// 
			/// **Returns:**
			/// 
			/// :``filepath``: `string` A full absolute filepath to the script to execute in 3ds Max.
			/// 
			/// **Author:**
			/// 
			/// Jeff Hanna, jeff.b.hanna@gmail.com, July 9, 2016 9:00:00 AM

			string filepath = args[args.Length - 1];
			if (filepath.StartsWith("-"))
			{
				string msg = String.Format("No script filepath provided. {0}", USAGE_INFO);
				show_message(msg);
			}
			else if (!System.IO.File.Exists(filepath))
			{
				string msg = String.Format("The specified script file does not exist on disk. {0}", USAGE_INFO);
				show_message(msg);
			}

			filepath = filepath.Replace("\\", "\\\\");
			return filepath;
		}


		static void Main(string[] args)
		{
			/// The main execution function of MXSPyCOM
			/// 
			/// **Arguments:**
			/// 
			/// :``args``: `string[]` The arguments provided to MXSPyCOM
			/// 
			/// **Keyword Arguments:**
			///
			/// None
			/// 
			/// **Returns:**
			/// 
			/// None
			/// 
			/// **Author:**
			/// 
			/// Jeff Hanna, jeff.b.hanna@gmail.com, July 9, 2016 9:00:00 AM
			
			if (args.Length == 0)
			{
				show_message("help");
			}
			else
			{
				string filepath = get_script_filepath(args);
				execute_max_commands(args, filepath);
			}

			Environment.Exit(0);
		}
	}
}
