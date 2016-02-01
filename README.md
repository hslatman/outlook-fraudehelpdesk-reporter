# Fraudehelpdesk Reporter Outlook Add-In 

An Outlook Add-In providing an easy way for forwarding suspicious emails as an attachment to the Fraudehelpdesk.

## Development

Visual Studio 2013 has been used to develop the plugin, and will most likely give the best results, although newer versions of Visual Studio should also work.

### Requirements

* Visual Studio (Tested with VS Ultimate 2013 and [2015](https://github.com/MSAdministrator/PhishReporter-Outlook-Add-In#requirements-to-buildcustomize-the-phishreporter-outlook-add-in))
* Visual Studio [Installer Projects Extension](https://visualstudiogallery.msdn.microsoft.com/9abe329c-9bba-44a1-be59-0fbf6151054d)
* Visual Studio Office Developer Tools

### Develop & Build

1. Clone the repository
2. Copy the ForwardToAbuseAddIn/app.config.dist file (Debug) and place it in the same folder, removing the .dist extension
3. Copy the ForwardToAbuseAddIn/pub/app.config.dist file (Release) and place it in the same folder, removing the .dist extension
4. Set the ReportAddress and ManifestUrl settings in both files to entities you control
  * There are two configuration files, one for Debug and one for Release builds
  * Visual Studio only knows about the Debug app.config, so you have to manually changed the Release app.config file
5. Most of the functional code resides in the ClickCode.vb module
  * Changed the other files to change more advanced functionality
6. Rebuild the FraudehelpdeskReporter project using Debug or Release configuration
7. Rebuild the Setup project
  * The Setup project always uses the files generated in the Release configuration (i.e. the .vsto and .dll)
  * It does use the right configuration file (i.e. Debug configuration for FraudehelpdeskReporter when building Setup with the Debug configuration)
  * Optionally, combine the setup files like described below, to create a single setup package
8. Deploy the fhd_download_manifest.xml file and the installer package to the location set in the corresponding app.config file
  * Make changes to the fhd_download_manifest.xml file: set the version of the assembly that was compiled and specify a location for the new version download
  * The fhd_download_manifest.xml file and the installer package do not necessarily have to be in the same place: it depends on the contents of the 
  fhd_download_manifest.xml file where the installer should be placed. Placing it on a secure location (HTTPS) is recommended!
  
  
## Combined Setup

After compilation of the Setup Project, there will be two files in the Setup/(Debug|Release) folder. These can be 
combined using IExpress. Simply starting IExpress from the commandline, and create a new package. Also see: 
http://stackoverflow.com/questions/535966/merge-msi-and-exe. This way the end user only has to download a single, 
self-extracting file, which can (optionally) start the setup.

The reason two files are created is because of the fact that the specification of prerequisites does not happen in the .msi file, but is done in the setup.exe file.
Without the setup.exe it is possible that the installation of the add-in works, but can't be used effectively in Outlook.
When setup.exe can't find the .msi file, the installation will fail for sure.

## Distribution

I put some research into what the best distribution method would be for an add-in like the FraudehelpdeskReporter.
Several options were considered, which are described below, including some pros and cons:

* Squirrel - Easy to create, easy to use, no support for VSTO, supports updates, no .NET 4.0 support, no XP support
* WiX - No auto-updates mechanism, slightly more complicated compared to Squirrel
* Omaha - Open sourced by Google and powers Google products like Chrome, supports auto updates but is quite complex and probably overkill for the project
* ClickOnce - Really simple to setup, natively supported by Visual Studio, VSTO and auto updates support, problems with installation (updating) behind proxies
* Windows Installer - Supported by Visual Studio, support for VSTO deployment, no auto updates supported

### Distribution Requirements

* Broad operating system support, i.e. Windows XP and up and both 32 and 64 bit versions
* Broad Microsoft Office version support, at least 2010, but the goal is to also support 2007 (again, both 32 and 64 bit support)
* Easy for the end user to install
* Should allow for code signing
* The add-in should have the option to be updated
* Options set by the user should of course be saved for future usage

Based on the requirements and the likely properties of the environments where the add-in will be deployed, the Windows Installer route was chosen, 
because it supports a broad variety of .NET framework versions, supports VSTO and includes the creation of locations to save settings. It also 
allows to install pre-requisites and can be packed in a single downloadable installer using IExpress.

### Distribution & Updates

Because the Windows Installer project does not support auto update functionality by default, I created a simple check for updates myself.
It is based on performing a lookup on an web address (i.e. _ManifestUrl_) where the _fhd_download_manifest.xml_ file resides. The _version_ attribute in the
_fhd_ node is read and compared to the version of the compiled assembly to see whether new versions are available. New versions shall be obtained via the
_url_ attribute. After the check for the update, the user is presented with a dialog to download the update, which is performed via the default browser.
The user can than download and install the new version of the add-in, which Windows will check for integrity.


## Installation

Windows binaries will be available soon!

## Contribute

Contributions are always welcome! Clone the project and create a PR as soon as your contribution is ready for review.

## Roadmap

Possible additions and changes:

* ~~Easier customization when developing~~
* ~~Customization options for end users~~
* ~~Show report button in additional places, like the email viewer~~
* ~~Cleaner (de)installation process if possible~~
* Add a post-build event that combines the setup.exe and .msi installer. (i.e. IExpress)
* Better localization, using resource files
* Clearer code documentation
* Refactoring

## Known Issues

* For some reason when changing the properties of Form2 (Options) the checkedState property of checkbox2 (Check for Updates) is sometimes set. 
  The _Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked_ line should be deleted.
* As described before, when adding new settings to the FraudehelpdeskReporter project, these have to be copied manually from the
  app.config to the pub/app.config file.


## License

Licensed under [GPLv3](LICENSE).

## Credits

This project is based on the work performed by [Josh Rickard](http://msadministrator.com/) as described in his blog post on http://msadministrator.com/2015/10/31/phishreporter-outlook-add-in/. Later an updated version appeared on [GitHub](https://github.com/MSAdministrator/PhishReporter-Outlook-Add-In).