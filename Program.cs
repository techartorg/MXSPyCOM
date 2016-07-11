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

		const string USAGE_INFO = "Type \"MXSPyCOM /?\" for usage info.";

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


		static void show_message(string message, bool exit = true)
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
			/// :``exit``: `bool` Controls whether or not MXSPyCOM should quit execution after the user has clicked OK on the message dialog.
			/// 
			/// **Returns:**
			/// 
			/// None
			/// 
			/// **Author:**
			/// 
			/// Jeff Hanna, jeff.b.hanna@gmail.com, July 9, 2016 9:00:00 AM

			MessageBoxIcon icon = MessageBoxIcon.Error;

			if (message.ToLower() == "help")
			{
				message = @"Used to execute MaxScript and Python scripts in 3ds Max.

Usage:
MXSPyCOM.exe [options] /f <filename>

Commands:
<filename>	- Full path to the script file to execute.

Options:
/debug		- Display debug output.
/o		- Echo output buffer.";
				icon = MessageBoxIcon.Information;
			}

			MessageBox.Show(message, "MXSPyCOM", MessageBoxButtons.OK, icon);

			if (exit)
			{
				Environment.Exit(0);
			}
		}


		static string validate_args(string[] args)
		{
			/// Validates the command line arguments used when MXSPyCOM was called.
			/// If no arguments are provided or the /? or /help arguments are provided then a help/usage dialog will be displayed.
			/// If /f is provided a full absolute filepath to the script file to execute is expected to be the next argument.
			/// If there is no filepath provided or if the file does not exist on disk an error dialog will be displayed.
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

			string msg = "";

			if (args.Length == 0 || Array.IndexOf(args, "/?") >= 0 || Array.IndexOf(args, "help") >= 0)
			{
				show_message("help");
			}

			int idx = Array.IndexOf(args, "/f");
			if (idx < 0)
			{
				msg = String.Format("No script filepath specified. {0}", USAGE_INFO);
				show_message(msg, false);
				return "";
			}
			else
			{
				try
				{
					string filepath = args[idx + 1];

					if (!System.IO.File.Exists(filepath))
					{
						msg = String.Format("The specified script file does not exist on disk. {0}", USAGE_INFO);
						show_message(msg);
					}

					string ext = System.IO.Path.GetExtension(filepath);
					if (ext.ToLower() == ".py")
					{
						filepath = make_python_wrapper(filepath);
					}

					return filepath;
				}
				catch(IndexOutOfRangeException)
				{
					msg = String.Format("No script file provided. {0}", USAGE_INFO);
					show_message(msg, true);
					return "";
				}
			}
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
				string filepath = "";
				//string[] options;

				filepath = validate_args(args);
				filepath = filepath.Replace("\\", "\\\\");

				if (filepath != "")
				{
					bool max_running = is_process_running("3dsmax");
					if (max_running)
					{
						var com_type = Type.GetTypeFromProgID("Max.Application");
						dynamic com_obj = Activator.CreateInstance(com_type);
						try
						{
							com_obj.filein(filepath);
						}
						catch (System.Runtime.InteropServices.COMException)
						{

						}
					}
					else
					{
						string msg = "3ds Max (3dsmax.exe) is not currently running.";
						show_message(msg);
					}
				}
			}

			Environment.Exit(0);
		}
	}
}
