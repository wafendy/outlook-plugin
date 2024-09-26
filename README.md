# Outlook Plugin Sample

# Prerequisite

Must purchase a license to enable "Add Custom Addins"

# Yeoman Generator

Setup Yeoman Generator by following the instructions here

https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview

# Macbook

Sideloading not working in OSX. Must run the dev server first, then add the manifest file manually.

```
- Install Microsoft Outlook
- Add profile (With licensed), if you have multiple profile ensure the licensed profile is the default profile
- Setup Yeoman Generator
- clone this repo
- npm install
- npm run dev-server
- manually add the manifest file (see screenshots below)
```

### console.log

```
# From terminal
defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true
```

### clearing cache

```
rm -fr ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef
rm -fr ~/Library/Containers/com.Microsoft.Outlook/Data/Library
rm -fr ~/Library/Containers/com.microsoft.Outlook/Data/Library/Application\ Support/Microsoft/Office/16.0/Wef/
rm -fr ~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/
rm -fr ~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/
```

# Windows

Sideloading works in Windows.

```
- Install Microsoft Outlook
- Add profile (With licensed), if you have multiple profile ensure the licensed profile is the default profile
- Setup Yeoman Generator
- clone this repo
- npm install
- npm run start (Sideloading)
```

# Screenshots

![outlook home](https://github.com/wafendy/outlook-plugin/raw/main/screenshots/home.png)

![outlook home](https://github.com/wafendy/outlook-plugin/raw/main/screenshots/office-add-ins.png)

![outlook home](https://github.com/wafendy/outlook-plugin/raw/main/screenshots/add-custom.png)

# Resources

ScriptLab: https://appsource.microsoft.com/en-us/product/office/WA200001603

office-js-snippets: https://github.com/OfficeDev/office-js-snippets

enable console log in Mac: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-office-add-ins-on-ipad-and-mac
