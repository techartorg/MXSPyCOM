(
	/* Dynamically writes the necessary Registry information to allow
		Simon Felton's MXSCOM bridge to work.
		IF RUNNING THIS SCRIPT ON AN VERSION OF MAX OLDER THAN MAX 10
		THE AVG EXTENSION *MUST* BE INSTALLED
	*/

	local reg_key
	local max_version = ((maxVersion())[1] / 1000) as string

	fn create_reg_key hkey key_name &reg_key key_value_name key_value_type key_value =
	(
		registry.createKey hkey key_name key:&reg_key
		registry.setValue reg_key key_value_name key_value_type key_value
	)

	fn write_sub_key_data reg_key sub_key_name sub_key_type sub_key_value =
	(
		local sub_key
		registry.createKey reg_key sub_key_name key:&sub_key
		registry.setValue sub_key "" sub_key_type sub_key_value
	)


	-- Establish a root key for generalized Max data
	create_reg_key HKEY_CURRENT_USER @"Software\Classes\MAX.Application" &reg_key "" #REG_SZ "OLE Automation MAX Application"

	-- Add the Clsid information
	write_sub_key_data reg_key "Clsid" #REG_SZ "{7FA22CB1-D26F-11d0-B260-00A0240CEEA3}"

	-- Add the CurVer information
	write_sub_key_data reg_key "CurVer" #REG_SZ ("MAX.Application." + max_version)

	-- Establish a new root key for the version of Max being used
	create_reg_key HKEY_CURRENT_USER (@"Software\Classes\MAX.Application." + max_version) &reg_key "" #REG_SZ ("OLE Automation MAX " + max_version + ".0 Application")

	-- Add the Clsid information
	write_sub_key_data reg_key "Clsid" #REG_SZ "{7FA22CB1-D26F-11d0-B260-00A0240CEEA3}"

	-- Make a new root key for the CLSID data
	create_reg_key HKEY_CURRENT_USER @"Software\Classes\CLSID\{7FA22CB1-D26F-11d0-B260-00A0240CEEA3}" &reg_key "" #REG_SZ ("OLE Automation MAX " + max_version + ".0 Application")

	-- Add sub key data
	write_sub_key_data reg_key "ProgID" #REG_SZ ("MAX.Application." + max_version)
	write_sub_key_data reg_key "VersionIndependentProgID" #REG_SZ "MAX.Application"

	-- Register the running of files and executing script code to OLE.
	registerOLEInterface #( filein, execute, edit, encryptscript )

)