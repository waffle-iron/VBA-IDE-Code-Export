# VBA IDE CodeExport
Export VBA code from the VBA IDE using the VBComponent class, currently this will export *.cls, *.bas, *.frm files to the root of the VBA project being exported.

## Description
It will create a menu in the VBA IDE called 'Export for VCS' with the options 'Make File List', 'Import Files', 'Export Files' and 'Configure Export'.

The default process is:

1. Make a File List
    - This will list all the objects in the project ina newly created module called 'modFileList'.
2. Export the Files
    - This will export the files listed in 'modFileList' to the root folder of the current project and remove all but the 'modFileList' module from the project.
3. Import the files
    - This uses the 'modFileList' module to build the project, all the files should be present to built the project.

Additional functionality is there to configure a .conf file that will export the project contents to that instead of the 'modFileList' leaving your project completely empty. This can also be configured to set the import and export file locations.

## Build
You will need to clone or download a zip of the project open an empty workbook in Excel then open the 'VBAIDECodeExport.xlam' add-in, you will be able to see the project from the VBA IDE [Alt+F11] pane of the new workbook you have just opened. The easiest way to do this is to drag and drop the components from file explorer to the VBA IDE

GIF HERE

Save and open again, the auto_open should take care of creating the VBA IDE menu options and you're good to go

## Still to come
The ability to specify file locations
The option to check that the file being worked on is in source control and if it is read only or not

## Contributing
Please fork this repository and contribute back using pull requests.

Any contributions, large or small, major features, bugfixes and integration tests are welcomed and appreciated but will be thoroughly reviewed and discussed.
