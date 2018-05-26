# MXSPyCOM
A modern version of MXSCOM, to allow for editing &amp; execution of 3ds Max MaxScript and Python files from external code editors.

In 2005 Simon Feltman released the first MXSCOM, a small Visual Basic 6 application that took commands and sent them to Autodesk's
3ds Max's internal COM server. This allowed users to choose their own external code editor for editing MaxScript and to be able to 
have their MaxScript code execute in 3ds Max by way of having the code editor utilize MXSCOM to send the file into 3ds Max and have it
executed. Modern versions of Windows can not use Simon Feltman's old MXSCOM.exe program due to it being ActiveX based.

MXSPyCOM is a C# based replacement for MXSCOM. It offers the same functionality as MXSCOM but can run on modern versions of Windows.
It also supports editing of Python files and having them execute in versions of 3ds Max, starting with 3ds Max 2015, that support Python
scripts.

Instructions on how to install MXSPyCOM and how to configure most popular code editors to use it as an external tool can be found in [this project's Wiki](https://github.com/JeffHanna/MXSPyCOM/wiki).
