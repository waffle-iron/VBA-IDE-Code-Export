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

1. Open the template file `VBA-IDE-Code-Export.xlsm`
2. Import the files specified in `CodeExportFileList.conf` (Tip: Use a previously installed copy of this Add-In)
3. Compile project as a smoke test
4. Save as an Add-In.

## Install

Save the Add-In in your Add-Ins folder. Add-Ins placed here will be loaded automatically. Once the Add-In is installed, enable the Add-In in Excel.

## Contributing
Please fork this repository and contribute back using pull requests.

Any contributions, large or small, major features, bugfixes and integration tests are welcomed and appreciated but will be thoroughly reviewed and discussed.

## Roadmap

- [] Add pretty ribbon UI
- [] Save XL as XML

