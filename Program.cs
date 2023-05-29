using System;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

using Westwind.Utilities;


namespace MXSPyCOM
{
    class Program
    {
        /// A modern version of MXSCOM, to allow for editing & execution of 3ds Max MaxScript and Python files from external code editors.
        /// In 2005 Simon Feltman released the first MXSCOM, a small Visual Basic 6 application that took commands and sent them to
        /// Autodesk's 3ds Max's internal COM server. This allowed users to choose their own external code editor for editing MaxScript
        /// and to be able to have their MaxScript code execute in 3ds Max by way of having the code editor utilize MXSCOM to send the file
        /// into 3ds Max and have it executed. Modern versions of Windows can not use Simon Feltman's old MXSCOM.exe program due
        /// to it being ActiveX based.
        ///
        /// MXSPyCOM is a C# based replacement for MXSCOM. It offers the same functionality as MXSCOM but can run on modern versions of Windows.
        /// It also supports editing of Python files and having them execute in versions of 3ds Max, starting with 3ds Max 2015, that support
        /// Python scripts.
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
        /// Jeff Hanna, jeff@techart.online, July 9, 2016

        const string USAGE_INFO = "\nType \"MXSPyCOM\" for usage info.";


        static int Get_MaxVersion(object com_obj)
        {
            var com_t = com_obj.GetType();
            var result = com_t.InvokeMember("execute", ReflectionUtils.MemberAccess | BindingFlags.InvokeMethod,
                                          null,
                                          com_obj,
                                          new object[] { "MaxVersion()" });
            if (result is object[] version_numbers)
            {
                return (int)version_numbers[0];
            }
            return 0;
        }

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
            /// Jeff Hanna, jeff@techart.online, July 11, 2016

            bool max_running = is_process_running("3dsmax");
            if (max_running)
            {
                if (args.Length == 1)
                {
                    string msg = string.Format("No options provided.", USAGE_INFO);
                    show_message(msg);
                }
                else
                {
                    string prog_id = "Max.Application";
                    Type com_type = Type.GetTypeFromProgID(prog_id);
                    object com_obj = Activator.CreateInstance(com_type);

                    string ext = Path.GetExtension(filepath).ToLower();

                    var version = Get_MaxVersion(com_obj);

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
                                    com_obj.GetType().InvokeMember("filein",
                                                                   ReflectionUtils.MemberAccess | BindingFlags.InvokeMethod,
                                                                   null,
                                                                   com_obj,
                                                                   new object[] { filepath });
                                }
                                catch (TargetInvocationException) { }
                                break;

                            case "-s":
                                try
                                {
                                    filepath = mxs_try_catch_errors_cmd(filepath, version);
                                    com_obj.GetType().InvokeMember("execute",
                                                                   ReflectionUtils.MemberAccess | BindingFlags.InvokeMethod,
                                                                   null,
                                                                   com_obj,
                                                                   new object[] { filepath });
                                }
                                catch (TargetInvocationException) { }
                                break;

                            case "-e":
                                try
                                {
                                    com_obj.GetType().InvokeMember("edit",
                                                                   ReflectionUtils.MemberAccess | BindingFlags.InvokeMethod,
                                                                   null,
                                                                   com_obj,
                                                                   new object[] { filepath });
                                }
                                catch (TargetInvocationException) { }
                                break;

                            case "-c":
                                if (ext == ".ms")
                                {
                                    try
                                    {
                                        com_obj.GetType().InvokeMember("encryptscript",
                                                                       ReflectionUtils.MemberAccess | BindingFlags.InvokeMethod,
                                                                       null,
                                                                       com_obj,
                                                                       new object[] { filepath });
                                    }
                                    catch (TargetInvocationException) { }
                                }
                                else
                                {
                                    string msg = $"Only MaxScript files can be encrypted. {USAGE_INFO}";
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
                string msg = "3ds Max is not currently running.";
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
            /// Jeff Hanna, jeff@techart.online, July 9, 2016

            Process[] pname = Process.GetProcessesByName(process_name);
            if (pname.Length > 0)
            {
                return true;
            }

            return false;
        }


        static string make_python_import_reload_wrapper(string python_filepath)
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
            /// :``reload_wrapper_filepath``: `string`
            ///
            /// **Author:**
            ///
            /// Jeff Hanna, jeff@techart.online, November 22, 2019

            string mod_name = Path.GetFileNameWithoutExtension(python_filepath);
            string reload_cmd = $"import sys\nif sys.version_info.major < 3:\n\timport imp as importlib\nelse:\n\timport importlib\n\t\ntry:\n\timportlib.reload({mod_name})\nexcept:\n\tpass";
            var reload_wrapper_filepath = Path.Combine(Path.GetTempPath(), "import_reload.py");
            reload_wrapper_filepath = reload_wrapper_filepath.Replace(@"\", @"\\");
            File.WriteAllText(reload_wrapper_filepath, reload_cmd);

            return reload_wrapper_filepath;
        }


        static string make_python_wrapper(string python_filepath)
        {
            /// For the python file being executed in 3ds Max this script writes a Maxscript wrapper file
            /// that can be called to reimport that Python module so that the in-memory version is updated with
            /// any changes made between script executions.
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
            /// Jeff Hanna, jeff@techart.online, July 9, 2016

            string reload_filepath = make_python_import_reload_wrapper(python_filepath);
            string cmd = $"python.ExecuteFile(\"{reload_filepath}\");python.ExecuteFile(\"{python_filepath}\")";
            var wrapper_filepath = Path.Combine(Path.GetTempPath(), "maxscript_python_wrapper.ms");
            File.WriteAllText(wrapper_filepath, cmd);

            return wrapper_filepath;
        }


        static string mxs_try_catch_errors_cmd(string filepath, int version)
        {
            /// Wraps the given MAXScript command arg in MAXScript code that when run in 3ds Max will catch
            /// errors using a MAXScript try() catch() and print a 'homemade' minimally useful log message
            /// to the MAXScript Listener instead of the usual proper one.
            ///
            /// **Arguments:**
            ///
            /// :``command: `string` A fumaxscript command to wrap
            /// :``log_filepath``: `string` The filepath to use in the MAXScript Listerner log message
            ///
            /// **Keyword Arguments:**
            ///
            /// None
            ///
            /// **Returns:**
            ///
            /// :``cmd``: `string`
            ///
            /// **Author:**
            ///
            /// Gary Tyler, mail@garytyler.com, April 12, 2018

            string ext = Path.GetExtension(filepath).ToLower();

            string location;
            string run_cmd;            
            
            if (ext == ".py")
            {
                /// For python files, use passed filepath for location msg, no pos or line available
                location = $"\"Error; filename: {filepath}\"";

                /// Pass thru python.ExecuteFile()
                string reload_filepath = make_python_import_reload_wrapper(filepath);
                run_cmd = $"python.ExecuteFile(\"{reload_filepath}\");python.ExecuteFile(\"{filepath}\")";
            }
            else
            {                                
                if (version > 19000 || version == 0)
                {
                    /// For maxscript files, use provided commands for location msg
                    string loc_file = "\" filename: \" + (getErrorSourceFileName() as string)";
                    string loc_pos = "\"; position: \" + ((getErrorSourceFileOffset() as integer) as string)";
                    string loc_line = "\"; line: \" + ((getErrorSourceFileLine() as integer) as string) + \"\n\"";
                    string callstack_line = "\"callstack: \n\" + (cs as string)";
                    location = $"#(\"Error;\" + {loc_file} + {loc_pos} + {loc_line} + {callstack_line})";
                }
                else
                {
                    string loc_file = "\" filename: \" + (getSourceFileName() as string) + \"\n\"";
                    string callstack_line = "\"callstack: \n\" + getCurrentExceptionStackTrace()";
                    location = $"#(\"Error;\" + {loc_file} + {callstack_line} + \"-- ########################################################################\n\" + getCurrentExceptionCallStack() + \"-- ########################################################################\n\")";                    
                }

                /// Pass thru filein()
                run_cmd = $"filein(@\"{filepath}\")";
            }

            string exception_array = "(filterString (getCurrentException()) \"\n\")";
            string exception_msg = $"(for i in {exception_array} + {location} do setListenerSelText (\"\n\" + i))";
            string cmd = $"try({run_cmd}) catch(cs = \"\" as stringStream;stack to:cs;{exception_msg});setListenerSelText \"\n\"";
            return cmd;
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
            /// Jeff Hanna, jeff@techart.online, July 9, 2016

            MessageBoxIcon icon = info == true ? MessageBoxIcon.Information : MessageBoxIcon.Error;

            if (message.ToLower() == "help")
            {
                message = @"Used to execute MaxScript and Python scripts in 3ds Max.

Usage:
MXSPyCOM.exe [options] <filename>

Options:
-f	- Execute the script in 3ds Max.
-s	- Execute the script in 3ds Max with no error dialogs.
-e	- Edit the script in 3ds Max's internal script editor.
-c	- Encrypt the script. Only works with MaxScript files.

Commands:
<filename>	- Full path to the script file to execute.";

                show_message(message, info: true);
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
            /// Jeff Hanna, jeff@techart.online, July 9, 2016

            string filepath = args[args.Length - 1];
            if (filepath.StartsWith("-"))
            {
                string msg = $"No script filepath provided. {USAGE_INFO}";
                show_message(msg);
            }
            else if (!File.Exists(filepath))
            {
                string msg = $"The specified script file does not exist on disk. {USAGE_INFO}";
                show_message(msg);
            }

            filepath = filepath.Replace(@"\", @"\\");
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
            /// Jeff Hanna, jeff@techart.online, July 9, 2016

            if (args.Length == 0)
            {
                show_message("help");

                // For testing
                //string filepath = @"d:\repos\mxspycom\hello_world.ms";
                //string[] test_args = new string[2] {"-f", filepath};
                //execute_max_commands(test_args, filepath);
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
