# VBA IDE CodeExport

[![The MIT License](https://img.shields.io/badge/license-MIT-orange.svg?style=flat-square)](http://opensource.org/licenses/MIT)

For a while now I have used this code so that all the associated VBA files used in a VBA project (*.cls, *.bas, *.frm files) can be easily exported for use with a Version Control System.

This is specifically for Excel, although the VBIDE extensibility can be used for all the MS Office suite.

## Installing

1. Obtain a copy of the add-in by following the build instructions below.
2. Save the add-in in your add-ins folder. Add-ins saved in your add-ins folder are loaded automatically.
3. Enable the add-in in Excel.
4. Check the `Trust access to the VBA project model` check box located in `Trust Centre -> Trust Centre Settings -> Macro Settings -> Trust access to the VBA project model`.

Optionally, set password protection to prevent the Add-In code annoying you in the VBE and to prevent accidental changes.

## Usage

The add-in will create a menu in the VBA IDE (the VBE) called `Export for VCS`. All controls for the add-in are found in this menu.

### The configuration file

A file named `CodeExport.config.json` in the same directory as an Excel file declares what gets imported into that Excel file. The `Make Config File` button in the `Export For VCS` menu will generate a new configuration file for the current active project based upon the contents of that project. Any existing configuration file will be overwritten. The JSON file format is used as the file format for the configuration file.

The `Module Paths` property specifies a mapping of VBA modules to their location in the file system. File paths may be either relative or absolute. Relatives paths are relative to the directory of the configuration file and the Excel file. The `Base Path` property can be used to add a common prefix to all the file paths.

The `References` property declares the references to libraries that your VBA modules require. These will be imported when the import action is used and will be removed when the export action is used.

The `VBAProject Name` property declares declares the VBAProject name. This will be imported when the import action is used.

### Importing

The `Import` button in the `Export For VCS` menu will:

* Import all the modules specified in the configuration file from the file system into the Excel file. Existing modules will be overwritten.
* Add all library references declared in the configuration file. Existing library references will be overwritten.
* Set the VBAProject name as declared in the configuration file.


### Exporting

The `Export` button in the `Export For VCS` menu will:

* Export all the modules specified in the configuration file from the Excel file into the appropriate places in the file system. Existing files will be overwritten.
* Remove library references from the project which are declared in the configuration file.

## Building

1. Open the template file `VBA-IDE-Code-Export.xlsm`.
2. Import the files specified in `CodeExport.config.json` (Tip: Use a previously installed copy of this Add-In).
3. Compile project as a smoke test.
5. Save as an Add-In.

## Contributing
Please fork this repository and contribute back using pull requests.

Any contributions, large or small, major features, bugfixes and integration tests are welcomed and appreciated but will be thoroughly reviewed and discussed.

Please use the template file `VBA-IDE-Code-Export.xlsm` for working in, however don't commit the template file unless you are actually making a change to the template file. This helps with source control since merging an Excel file is not fun.

## Roadmap

- [ ] Add pretty ribbon UI
- [ ] Save XL as XML
