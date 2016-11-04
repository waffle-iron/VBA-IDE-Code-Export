# VBA IDE CodeExport
For a while now I have used this code so that all the associated VBA files used in a VBA project (*.cls, *.bas, *.frm files) can be easily exported for use with a Version Control System.

This is specifically for Excel, although the VBIDE extensibility can be used fro all the MS Office suite.

## Description
On opening it will create a menu in the VBA IDE called 'Export for VCS' with the options 'Make File List', 'Import Files', 'Export Files' and 'Configure Export'.

The default process for use is:

1. Make a File List
    - This will list all the objects in the project in a newly created module called 'modFileList'.
2. Export the Files
    - This will export the files listed in 'modFileList' to the root folder of the current project and remove all but the 'modFileList' module from the project.
3. Import the files
    - This uses the 'modFileList' module to build the project, all the files should be present to built the project.

Additional functionality is there to configure a .conf file that will export the project contents to that instead of the 'modFileList' leaving your project completely empty. This can also be configured to set the import and export file locations.

## Build
You will need to clone or download a zip of the project open an empty workbook in Excel then open the 'VBAIDECodeExport.xlam' add-in, you will be able to see the project from the VBA IDE [Alt+F11] pane of the new workbook you have just opened. The easiest way to do this is to drag and drop the components from file explorer to the VBA IDE

GIF HERE

Save and open again, the auto_open should take care of creating the VBA IDE menu options and you're good to go

## Contributing
Please fork this repository and contribute back using pull requests.

Any contributions, large or small, major features, bugfixes and integration tests are welcomed and appreciated but will be thoroughly reviewed and discussed.
