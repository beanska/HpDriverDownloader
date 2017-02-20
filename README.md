#  HPDriverDownloader 0.2

Downloads HP Softpaqs and extracts them by OS / Model. Created as an alternative to the HP's Softpaq manager to organize softpaqs in a way that is easier to import in MDT/SCCM. 

Requires
	Powershell 5.0
	PS FTP Module
	Elevated Privleges (To extract Softpaqs)
	
Usage
	HPDriverDownloader.ps1 [-Config config.xml] [-OutDir D:\OutputDir] [-Update] [-SymLink]
	
	-Config		Specify a configuration file. (Defaults to script folder)
	-OutDir		Specify a directory to store Softpaqs. (Defaults to script folder)
	-Update 	Downloads a new HP catalog, updates config file.
	-SymLink	Creates symbolic links instead of creating muliple copies of Softpaq files. Saves large amounts of space. If not using -SymLink it is recommended to use deduplication.

Examples

	Initial Usage - You must run this to generate the config file.
	Don't forget to edit config file after running this
	**HPDriverDownloader.ps1 -OutDir D:\Drivers -Update**
	
	Downloads softpaq file to 'D:\Drivers'. Creates symlinks instead of copying files for each model.
	**HPDriverDownloader.ps1 -OutDir D:\Drivers -SymLink**
	
	
