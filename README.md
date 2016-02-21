Required tools
------------------------------------
1. Visual Studio code http://code.visualstudio.com
2. NodeJS https://nodejs.org/en/
3. Git for windows http://www.git-scm.com/downloads
4. GitHub user name

Setup
------------------------------------

1. Clone repo

`git clone https://github.com/jthake/PowerPoint-Add-In-ConvoTickr.git  `

2. In commmand prompt run NPM INSTALL to get all the packages

`npm install`

To run sample
-------------------------------------

1. In command prompt run node web server

`gulp serve-static`

2. Open up web browser and open up OneDrive from App Launcher e.g.

`https://onedrive.live.com/`

3. Click 'New' and 'PowerPoint Presentation'
4. in PowerPoint Online click 'Insert' in Ribbon and select 'Office add-ins'
5. Click on drop down arrow next to 'Manage my add-ins'
6. Select "Upload manifest" and select 'manifest-convotickr.xml'
7. Insert the add-in into the PowerPoint document