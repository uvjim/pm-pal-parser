# pm-pal-parser
This script is an extension of a piece of work started by a former colleague.  The idea behind the process is to show value for Ivanti Performance Manager.  This is achieved by a using a combination of software: -  

* Perfmon - https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc749154(v%3dws.11)
* PAL Parser - https://github.com/clinthuffman/PAL

## The Process
**_N.B._** _It is important that the machines being used in testing are as close to the same specification as possible.  It is also important to ensure that the same or extremely similar workloads are taking place as well.  Results will be skewed if there is too much deviation.  You will need some machines with Performance Manager installed and some without for reference purposes._  

1. Open Perfmon
2. Create a new Data Collector Set using the provided Data Collector Set template
3. Run the collection long enough to capture enough information to demonstrate normal usage - can be started based on a schedule
4. Stop the logging – can be stopped on a schedule
5.	Take the `.blg` file produced and run it through PAL using the attached threshold file (the threshold file must be in the same location as the `PALWizard.exe`)

The steps above need to be completed for each of the machines in questions - those with Performance Manager installed and those without.

You should now have a set of `.htm` files for machines with and without Performance Manager installed.  Ensure that you store these `.htm` files in a folder structure that represents this, e.g.

![alt text](https://github.com/uvjim/pm-pal-parser/blob/master/folder%20structure.jpg)

### What to do with the output
Now that you have that folder structure with the relevant `.htm` files in both folders, you'll need to run the script against them.
The script is used as follows: -  

`pm-pal-parser.ps1 -WithoutPMFolder "C:\PM PAL Parser\Without PM\" -WithPMFolder " C:\PM PAL Parser\With PM\" -Filename " C:\PM PAL Parser\CompareWithandWithout.xlsx" -Application "iexplore","chrome"`  

This example will run a comparison for each machine in the set and produce a general summary as well as something more targeted towards both IE and Chrome.

### Available Parameters
`WithoutPMFolder` - path to the folder containing the `.htm` files for those machines that *did not* have Performance Manager installed.  
`WithPMFolder` - path to the folder containing the `.htm` files for those machines that *did* have Performance Manager installed.  
`Filename` - the path to the output file.  
`Application` - a list of applications that you may be interested in - this gives the ability to report on just these processes.  
`TreatAsOne` - supplying this paramters will collpase all servers and treat them as if they are just one.  This can be useful when working with a set of machines (with and without PM installed) that have the same specifications and may be providing services from a virtual or provisioned environment.  
`Label` - this is the label that should be used when the `TreatAsOne` parameter is supplied.  Default is "All Endpoints".
