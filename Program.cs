using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace MXSPyCOM
{
	class Program
	{
		/// <summary>
		/// 
		/// </summary>
		/// 
		const string USAGE_INFO = "Type \"MXSPyCOM /?\" for usage info.";

		static bool is_process_running(string process_name)
		{
			/// <summary>
			/// 
			/// </summary>
			/// 

			Process[] pname = Process.GetProcessesByName(process_name);
			if (pname.Length > 0)
			{
				return true;
			}

			return false;
		}


		static string make_python_wrapper(string python_filepath)
		{
			/// <summary>
			/// 
			/// </summary>
			/// 

			string cmd = String.Format("python.ExecuteFile(@\"{0}\")", python_filepath);
			string WRAPPER_FILEPATH = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "maxscript_python_wrapper.ms");
			System.IO.File.WriteAllText(WRAPPER_FILEPATH, cmd);

			return WRAPPER_FILEPATH;
		}


		static void show_message(string message, bool exit = true)
		{
			/// <summary>
			/// 
			/// </summary>
			/// 

			MessageBoxIcon icon = MessageBoxIcon.Error;

			if (message.ToLower() == "help")
			{
				message = @"Used to execute MaxScript and Python scripts in 3ds Max.

Usage:
MXSPyCOM.exe [options] -f <filename>

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
			/// <summary>
			/// 
			/// </summary>
			/// 

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
				catch(IndexOutOfRangeException e)
				{
					msg = String.Format("No script file provided. {0}", USAGE_INFO);
					show_message(msg, true);
					return "";
				}
			}
		}


		static void Main(string[] args)
		{
			/// <summary>
			/// 
			/// </summary>
			/// 

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
