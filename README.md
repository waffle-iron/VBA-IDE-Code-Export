# VBAIDECodeExport
Export VBA code from the VBA IDE using the VBComponent class, currently this will export *.cls, *.bas, *.frm files to the root of the VBA project being exported.

Special thanks to Paul Crook for the base of this code! You the man Crookie xx

# Description
It will create a menu in the VBA IDE called 'Export for TFS' with the options 'Make File List', 'Import Files' and 'Export Files'

The process to use this should be:

    1. Make a File List
      a. This will list all the object in the project
    2. Export the Files
      b. This will export the files to the root folder of the current project and remove all but the modFileList from the project
    3. Import the files
      c. This uses the modFileList to build the project, all the files should be present to built the project

# Build
To build you will need an empty *.xlam [Microsoft add-in file] to add the 'modImportExport.bas' , 'menuModule.bas' and 'clsVBECmdHandler.cls' to and in the References options add a reference to the Microsoft Visual Basic for Applications Extensibility to it. Save and open again, the auto_open should take care of creating the VBA IDE menu options and you're good to go

# Still to come
The ability to specify file locations
The option to check that the file being worked on is in source control and if it is read only or not


# Contributing

Please fork this repository and contribute back using pull requests.

<<<<<<< HEAD
Any contributions, large or small, major features, bugfixes and integration tests are welcomed and appreciated but will be thoroughly reviewed and discussed.
=======
Any contributions, large or small, major features, bugfixes and integration tests are welcomed and appreciated but will be thoroughly reviewed and discussed.
>>>>>>> refs/remotes/origin/master
