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

3. In command prompt run NPM INSTALL to get gulp globally

`npm install -g gulp`

To run sample
-------------------------------------

1. In command prompt run node web server

`gulp serve-static`

2. Open up a web browser and open up the contents of the add-in running locally on your machine and trust the locally signed SSL web site:

`https://localhost:8443/app/home/home.html`

3. Open up web browser and open up OneDrive from App Launcher e.g.

`https://appdev365.sharepoint.com/_layouts/15/WopiFrame.aspx?sourcedoc={f5fb2758-9447-4e9e-beba-8259fbb84d6a}&action=editnew&Source=https%3A%2F%2Fappdev365%2Esharepoint%2Ecom%2FSitePages%2FHome%2Easpx`

Confirm that the add-in is running.