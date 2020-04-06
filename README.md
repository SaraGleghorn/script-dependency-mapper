# Script Dependency Mapper
This is a tool for unpicking your SQL scripts if you've gotten into a mess with many SQL scripts and you don't know what you can safely change or delete.
This crawls through all your SQL scripts to extract and map dependencies between tables and scripts, and returns all the tables used and what they're used for (e.g myTable is CREATED in script1.sql and SELECTED in script2.sql, so you know that if you delete script1.sql it'll break your script2.sql. Hopefully, your naming conventions are better than this.)

## Getting Started

### Prerequisites

- Microsoft Access
- Any number of .sql files that require their dependencies logged.

## Running the dependency mapper

Download ScriptDependencyMapper.accdb

Enable Macros
```
Note: As best practice security, I recommend that you or someone fluent in code actually read the code in the forms and VBA modules before enabling macros, to ensure that you are happy they do not contain any security risks.
```
On the main form, click on "Set Search Folders" and list the folders that contain the .sql scripts you want to get the dependencies for.

If you also want to search subfolders, tick the box in the "subfolders" column.

Close the table.

On the main form, click on "Find .SQL files." This will attempt to find all .sql files within the folders you listed (and their subfolders, if ticked). A popup will tell you how many were found.

On the main form, click "Detect dependencies." This will list all table references within your .sql files.

## Built With

Microsoft Access VBA

## Future Plans
I'll be working on generating a visual script hierarchy so that it's easier to see how scripts are related.

## Authors

* **Sara Gleghorn** - *Initial work* - [Sara Gleghorn](https://github.com/SaraGleghorn)

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
