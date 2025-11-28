---
title: Implement event-based activation in WXP add-ins 
description: Learn how to develop a WXP add-in that implements event-based activation.
ms.date: 11/27/2025
ms.topic: how-to
ms.localizationpriority: medium
---

# Implement event-based activation in WXP add-ins 

With the feature, a central deployed Word, Excel or PowerPoint add-in can launch automatically in the background whenever a document is created or opened. This allows the add-in to validate, insert, or refresh critical content without any manual user operations.

## Supported events and clients

| Event name | Description | Supported clients and channels |
| ----- | ----- | ----- |
| `OnDocumentSave` | Occurs on a user saves a document manually in WXP.| <ul><li> Office Win32 Desktop DevMain channel insider ring, version>= 16.0.19426.20094 </li></ul>|
| `OnDocumentSaveAs` | Occurs on a user saves a copy of document manually in WXP.| <ul><li> Office Win32 Desktop DevMain channel insider ring, version>= 16.0.19426.20094 </li></ul>|

The following sections walk you through how to develop a Word add-in that automatically checks the document when a new or existing document saves. This highlights a sample scenario of how you can implement event-based activation in WXP add-ins.

## Set up your environment

To run the feature, you must have a supported version of Word and a Microsoft 365 subscription. Then, create a Word add-in project. You can create an add-in by following [Word add-in quick start](../quickstarts/word-quickstart-yo.md) and try to create an Office Add-in Task Pane project or other.

## Configure the manifest

 Currently, only Add-in only manifest supported. To enable an event-based add-in in WXP, you must configure the following elements in the `VersionOverridesV1_0` node of the manifest.

- In the [Runtimes](/javascript/api/manifest/runtimes) element, override the using runtime with a javascript type and reference a javascript file containing the function you want to execute.
- Set the `xsi:type` of the [ExtensionPoint](/javascript/api/manifest/extensionpoint) element to [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent). This enables the event-based activation feature in your WXP add-in.
- In the [LaunchEvent](/javascript/api/manifest/launchevent) element, set the `Type` to `OnDocumentSave` `OnDocumentSaveAs` and specify the JavaScript function name of the event handler in the `FunctionName` attribute.

### Code sample

1. In your code editor, open the quick start project you created.
1. Open the **manifest.xml** file located at the root of your project.
1. Select the entire **\<VersionOverrides\>** node (including the open and close tags) and replace it with the following XML. (Word version)

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <Runtimes>
          <Runtime resid="Taskpane.Url" lifetime="long" />
          <Runtime resid="WebViewRuntime.Url">
            <Override type="javascript" resid="JsRuntimeWord.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnDocumentSave" FunctionName="checkParagraphOnSave"></LaunchEvent>
              <LaunchEvent Type="OnDocumentSaveAs" FunctionName="checkParagraphOnSave"></LaunchEvent>
            </LaunchEvents>
            <SourceLocation resid="WebViewRuntime.Url"/>
          </ExtensionPoint>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="CommandsGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="TaskpaneButton">
                  <Label resid="TaskpaneButton.Label"/>
                  <Supertip>
                    <Title resid="TaskpaneButton.Label"/>
                    <Description resid="TaskpaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
        <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="JsRuntimeWord.Url" DefaultValue="https://raw.githubusercontent.com/yilin4/AddinForDLP/refs/heads/main/src/commands/autoruncommandsWord.js"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
        <bt:String id="CommandsGroup.Label" DefaultValue="Event-based add-in activation"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="My add-in"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Your sample add-in loaded successfully. Go to the HOME tab and click the 'Show Task Pane' button to get started."/>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane"/>
      </bt:LongStrings>
    </Resources>
</VersionOverrides>
```

4. Save your changes.

## Implement the event handler

To enable your add-in to complete tasks when the `OnDocumentSave` event occurs, you must implement a JavaScript event handler. In this section, you'll create the `changeHeader` function that adds header of public or high confidential to a document when open it according to whether it's a new document or an old one that already has content.

1. From the same quick start project, navigate to the **./src/commands** directory.
1. In the **./src/commands** folder, create a new file named **autoruncommandsWord.js**.
1. Open the **autoruncommandsWord.js** file you created and add the following JavaScript code.

```javascript
/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

async function checkParagraphOnSave(event) {
  let allow = true;
  await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.load("text");
    await context.sync();
    if (paragraph.text.includes("123456")){
      allow = false;
    }
  });

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  if (allow) {
    event.completed({ allowEvent: allow });
  } else {
    event.completed({
      allowEvent: allow,
      errorMessage: "Do not include 123456!",
    });
  }
}
async function registerOnParagraphChanged(event) {
  Word.run(async (context) => {
    let eventContext = context.document.onParagraphChanged.add(paragraphChanged);
    await context.sync();
  });
  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope

Office.actions.associate("changeHeader", changeHeader);
Office.actions.associate("checkParagraphOnSave", checkParagraphOnSave);
```

4. Save your changes. In the manifest, replace the following content to your own url.
```xml
<bt:Url id="JsRuntimeWord.Url" DefaultValue="https://raw.githubusercontent.com/yilin4/AddinForDLP/refs/heads/main/src/commands/autoruncommandsWord.js"/>
```

## Add a reference to the event-handling JavaScript file

Ensure that the **autoruncommandsWord.js** file must be a javascript file not a typescript file, and the online url is recommended.

## Test and validate your add-in

1. Sideload your add-in in [Word online](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) or [Windows desktop](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).
1. Create a new Word document or open an existing one, and you will see the headers are added to the document.

## Deploy your add-in
Event-based add-ins work only when deployed by an administrator; if users install them directly from AppSource or the Office Store, the will not automatically launch. Admin deployments are done by uploading the manifest to the Microsoft 365 admin center. In the admin portal, expand the Settings section in the navigation pane then select Integrated apps. On the Integrated apps page, choose the Upload custom apps action. For more information about how to deploy an add-in, please refer to this [link](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/office-addins?view=o365-worldwide). 

## Behavior and limitations

As you develop an event-based add-in for WXP, be mindful of the following feature behaviors and limitations.
- The feature is only supported in add-in only manifest.
- Office MAC Desktop is not supported yet.
- If a user installs multiple add-ins with the same ativation event, only one add-in will be activated randomly.
- APIs with UI and Office common APIs are not supported.
