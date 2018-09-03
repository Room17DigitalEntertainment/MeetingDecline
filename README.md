# Meeting Decline ![logo](Room17DE.MeetingDecline/Resources/MeetingDeclineImage.png "Logo")

## Table of Contents
- [About](#about)
- [Features](#features)
- [Install](#install)
- [Usage](#usage)
- [Submit an issue](#submit-an-issue)
- [Development](#development)
  - [Contributing](#contributing)
- [License](#license)

### About
Outlook addin for automatically declining meeting requests received in certain folders.

### Features
- decline a meeting received in a folder
- or say tentative to it
- and send optional response with an optional message

### Install
The installation process is pretty straightforward: download latest installer from: [GitHub](https://github.com/Room17DigitalEntertainment/MeetingDecline/releases/latest), run `Setup.exe` then click `Install` if prompted. That's it!

> _Note: on multiple users system, the application must be installed by each user who wants to use it._

To uninstall, go to `Control Panel` -> `Programs` -> `Uninstall a program` and uninstall `Meeting Decline`.

### Usage
- Right click on any folder and select `Meeting Decline`

  ![main](screenshots/main.png?raw=true "main")
  
- Click on any column header to sort its contents

  ![sort](screenshots/sort.png?raw=true "sort")

- Hover your mouse over the folder name to see it's full path:

  ![hover](screenshots/hover.png?raw=true "hover")

- Check `Enabled` box to start declining email:

  ![enable](screenshots/enable.png?raw=true "enable")

- Optional, choose what action needs to be taken when receiving the meeting, between `Decline` or `Tentative`:

  ![choose](screenshots/choose.png?raw=true "choose")

- Optional, check `Send response` to send a response to the Meeting organizer, if he wanted to:

  ![response](screenshots/response.png?raw=true "response")

- Optional, click on `Message` link to add a custom message for response message:

    ![message](screenshots/message.png?raw=true "message")

    - A new window will appear where you can type in your message:

    ![input](screenshots/input.png?raw=true "input")

    - Click `OK` to save it, or `Cancel` to revert changes.

- That's it! The application will start listening to new emails and decline them according to your rules.

  ![finish](screenshots/finish.png?raw=true "finish")
  
### Submit an issue
Please use [GitHub issues](https://github.com/Room17DigitalEntertainment/MeetingDecline/issues) to add a new bug or feature request.

When opening a bug, please include information about which **Operating System** you're using, **Outlook version**, any error you've seen and what you were doing inside the application. Before that, please take a moment before submitting and check if there isn't an already existing bug opened. If so, then kindly add a comment describing your situation too.
  
### Development
Software Required:
- Visual Studio (minimum 2015, recomended 2017)
- Visual Studio Tools for Office (VSTO) - see in Visual Studio Installer
- .NET Framework >=4.5.2
- Microsoft Outlook

Initial build required a _non_ existent `.pfx` and `.snk` file in project directory to sign the assembly. Go to Project `Properties` -> `Signing`:
- select any option between `Select from Store...`, `Select from File...` or `Create Test Certificate...` to add another `.pfx` file
- untick `Sign the assembly` or generate a new `.snk` file under `Choose a strong name key file` -> `<New...>`

While debugging, Visual Studio will install the addin, start a new instance of Outlook.exe and attach the debugger to it. You may find usefull to set the environment variable `VSTO_SUPPRESSDISPLAYALERTS` to `0` so Outlook will report any uncaught exception from addin via a popup.

Use these steps in case Outlook doesn't want to load the addin:
- go to `File` -> `Options` -> `Add-ins`
  - if the addin is listed `Inactive Application Add-ins`
    - select `COM Add-ins` under `Manage:` (lower page section) then click `Go...`
    - check addin name and click OK
  - if the addin is listed under `Disabled Application Add-ins`
    - select `Disabled Items` under `Manage:` (lower page section) then click `Go...`
    - select the addin and click Enable
- restart Outlook

#### Contributing
Want to add something awesome to this project? We would love to see contributions from you! But first, slow down a bit and please make sure you follow these _common sense_ rules:
- Fork the repository to **your** GitHub account, so all changes happen in your account. Create a new _feature_ branch from `master` and apply your changes there. After that, simply create a `Pull Request` from your feature branch against `MeetingDecline` `master` branch.
- Please limit each `Pull Request` to **one** feature. Don't add many features/changes, like a _fix_ and a _UI change_, into a single `Pull Request`. Create one at a time.
- If you have complex changes, across multiple commits, try to `squash` them into one single commit. For instance, don't create a `Pull Request` consisting of 10 commits showing all your work in progress. Before you create the `Pull Request`, [squash](https://git-scm.com/book/en/v2/Git-Tools-Rewriting-History) your changes into a single commit.


### License
This project is licensed under _GNU General Public License v3.0_. See [LICENSE.txt](LICENSE.txt) for terms of the license.
