Access Database Cleanup
-----------------------

Snippets of code that help in the process of cleaning up a legacy Microsoft 
Access database. It allows the user to search for all references to a particular 
object, making it easier to determine what else will be affected by removing or 
altering an object.

**Exporting Database Objects**

*export.bas* contains a VBA subroutine to export all database objects to text 
files. Import this file or copy it into a module in the Access database being 
worked on. Update the *exportLocation* variable to the desired directory to 
export text files to.

**Searching For Object References**

*search.py* contains Python code to search exported text files for references to 
a given object. Update the *directory* variable to the output directory chosen 
in *export.bas* and update the *object_name* variable with the particular Access 
object to be removed or altered.