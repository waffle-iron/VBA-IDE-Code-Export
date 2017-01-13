# VBA IDE CodeExport

[![The MIT License](https://img.shields.io/badge/license-MIT-orange.svg?style=flat-square)](http://opensource.org/licenses/MIT)

For a while now I have used this code so that all the associated VBA files used in a VBA project (*.cls, *.bas, *.frm files) can be easily exported for use with a Version Control System.

This is specifically for Excel, although the VBIDE extensibility can be used for all the MS Office suite.

## Description
On opening it will create a menu in the VBA IDE called 'Export for VCS' with the options 'Make File List', 'Import Files', 'Export Files'.

The default process for use is:

1. Make a File List
    - This will list all the objects in the project in a newly created file called 'CodeExportFileList.conf' in the same directory as the VBA project file.
2. Export the Files
    - This will export the files listed in 'CodeExportFileList.conf' to the directory which contains the VBA project file and remove all modules from the project.
3. Import the files
    - This uses the 'CodeExportFileList.conf' file to build the project, all the files should be present to built the project.

## Build

1. Open Excel, and create a new blank workbook.
2. Open the VBE and import the following VBA source files:
    * `clsVBECmdHandler.cls`
    * `modImportExport.bas`
    * `modMenu.bas`
3. Create a reference to the following libraries:
    * `Microsoft Visual Basic For Applications Extensibility 5.3`
    * `Microsoft Scripting Runtime`
4. Save the workbook as an Excel Add-in.

## Contributing
Please fork this repository and contribute back using pull requests.

Any contributions, large or small, major features, bugfixes and integration tests are welcomed and appreciated but will be thoroughly reviewed and discussed.

## Roadmap

- [] Add pretty ribbon UI
- [] Save XL as XML

