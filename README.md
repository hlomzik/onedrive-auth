# onedrive-auth

Simple javascript OneDrive auth library. Makes an authorization and returns back the token.

Here is the simple example of authorizing using this library and interacting with OneDrive api to create a file explorer web app: [See it in action!](https://hlomzik.github.io/onedrive-auth)

Included in this project:

* OneDriveExplorer (index.html) - A sample web app to view the contents of the signed in user's OneDrive and show the JSON structures returned by the API.
* OneDriveAuth (odauth.js) - A simple js library for handling the OAuth2 implicit grant flow for OneDrive. Used by the OneDriveExplorer web app.
