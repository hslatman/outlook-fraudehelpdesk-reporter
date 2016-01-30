# Fraudehelpdesk Reporter Outlook Add-In 

An Outlook Add-In providing an easy way for forwarding suspicious emails as an attachment to the Fraudehelpdesk.

## Development

* Visual Studio (Tested with VS Ultimate 2013 and [2015](https://github.com/MSAdministrator/PhishReporter-Outlook-Add-In#requirements-to-buildcustomize-the-phishreporter-outlook-add-in))
* Visual Studio Office Developer Tools
* Visual Studio [Installer Projects Extension](https://visualstudiogallery.msdn.microsoft.com/9abe329c-9bba-44a1-be59-0fbf6151054d)

1. Clone the repository
2. Make changes to Ribbon1.vb
  * More advanced changes can be made in other files
3. Rebuild the FraudehelpdeskReporter project
4. Rebuild the Setup project


## Combined Setup

After compilation of the Setup Project, there will be two files in the Setup/(Debug|Release) folder. These can be 
combined using IExpress. Simply starting IExpress from the commandline, and create a new package. Also see: 
http://stackoverflow.com/questions/535966/merge-msi-and-exe


## Installation

Windows binaries will be available soon!

## Contribute

Contributions are always welcome! Clone the project and create a PR as soon as your contribution is ready for review.

## Roadmap

Possible additions and changes:

* Easier customization when developing
* Customization options for end users
* ~~Show report button in additional places, like the email viewer~~
* ~~Cleaner (de)installation process if possible~~
* Add a post-build event that combines the setup.exe and .msi installer. (i.e. IExpress)

## License

Licensed under [GPLv3](LICENSE).

## Credits

This project is based on the work performed by [Josh Rickard](http://msadministrator.com/) as described in his blog post on http://msadministrator.com/2015/10/31/phishreporter-outlook-add-in/. Later an updated version appeared on [GitHub](https://github.com/MSAdministrator/PhishReporter-Outlook-Add-In).