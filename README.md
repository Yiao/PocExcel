# Office-Addin-TaskPane-Angular

1. You need install packages with npm 
```
npm install
```

2. when it ask you to add certificat, select "yes" 

3. Start application with npm
```
npm start
```
when you get this error with npm start
```
> office-addin-taskpane-js@0.0.1 start E:\My Office Add-in
> office-addin-debugging start manifest.xml

Debugging is being started...
App type: desktop
Enabled debugging for add-in af71531f-d6c5-4ff3-854d-7b8dc52e3dcd. Debug method: 0
Starting the dev server... (webpack-dev-server --mode development)
Unable to start the dev server. Error: The dev server is not running on port 3000.
Sideloading the Office Add-in...
Debugging started.
```

just run with
```
sudo npm install
```

4. Run back service [PocExcelBackService](https://github.com/Yiao/PocExcelBackService)
Fellow the readme.md in [PocExcelBackService](https://github.com/Yiao/PocExcelBackService/blob/main/README.md)

5. debug : https://github.com/OfficeDev/office-js-docs-pr/blob/master/docs/testing/debug-office-add-ins-on-ipad-and-mac.md

6. sideload : [How to sideload application](https://docs.microsoft.com/fr-fr/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)

This repository contains the source code used by the [Yo Office generator](https://github.com/OfficeDev/generator-office) when you create a new Office Add-in that appears in the task pane. You can also use this repository as a sample to base your own project from if you choose not to use the generator. 

## TypeScript

This template is written using [TypeScript](http://www.typescriptlang.org/). For the JavaScript version of this template, go to [Office-Addin-TaskPane-Angular-JS](https://github.com/OfficeDev/Office-Addin-TaskPane-Angular-JS).

## Debugging
when you got run issus, you could try to run thoese commandes:
```
npm uninstall - g office-addin-debugging
npm install - g office-addin-debugging
```

This template supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API).  If your question is about the Office JavaScript APIs, make sure it's tagged with  [office-js].

## Additional resources

* [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* More Office Add-in samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

Share Test
* [Solution Shared-folder-catalog](https://docs.microsoft.com/en-US/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins#sideload-your-add-in)

## Copyright

Copyright (c) 2019 Microsoft Corporation. All rights reserved.
