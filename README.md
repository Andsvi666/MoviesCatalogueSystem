# Movie Catalogues System

System to catalogue movies by making .xlsx files. Functionality is used all within console and project is fully made using C#

## Description

This is a simple system for cataloguing various films. Each newly created catalogue is saved as an .xlsx file. Additionally, there are two preset files, "Watchlist" and "Watched Movies", which track movies you want to see in the future and movies you have already watched. Films can be added to catalogues by entering all relevant information, and they can also be removed. The system includes functionality to generate automated catalogues based on selected criteria: it collects movies matching the criteria from all existing catalogues and creates a new catalogue automatically.

### Dependencies

* .NET Runtime (or .NET SDK if running from source)

### Installing

* Donwload project from github
* Unzip the folder in desired location

### Executing program

* Within main folder click "run.bat" to run the project

Or

* Run the terminal
* Select project folder
* Write the following command
```
dotnet run
```

## Help

All commands in the game must be typed exactly as shown. For example, to create an automated catalogue of all 1940s movies, press 7 to select the automated catalogue option, then press a to choose a year or decade. Next, press d to select a decade and type the decade in full, such as 1940. Command input is case-sensitive and must match exactly; if a mistake is made at any step, the program will reset back to the main menu.
