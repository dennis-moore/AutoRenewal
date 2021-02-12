# AutoRenewal
Automatically prepare renewals for legal review

### Config files
Json configuration files are used to control the mapping of data from excel to word. Currently only the PS1 organization is configured. In order to add support for new organizations, a new json config file can be created with the desired mappings, using the PS1 config as a guide. 

On application start, the files inside the configs directory are parsed and added to the organization dropdown list. This means that if new configs are added while the app is running, the app will need to be restarted before it sees the new configs.

### Future work
1) Create a user control to model each organization type and hold state information for each run
2) Add support for "batch" runs - 2 or more input files can be included in a run
3) Allow for batching of different organizations - i.e. two PS1's and one PS2 run at the same time
4) Add MVVM
5) Improved error handling
6) Adding / modifying json configs / organizations via the app

### Known issues
1) Word template files with complex headers and footers corrupt the output file and prevent it from being opened with Microsoft Word.
