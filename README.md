# onedrive-auth

Simple javascript OneDrive auth library. Makes an authorization and returns back the token.

## Usage

```
bower install onedrive-auth
```

Use the main file (`dist/odauth.js`) in any way you prefer: `require()` it or
directly add via `<script>` tag and use global `OneDriveAuth` object:

```js
// <script src="js/odauth.js"> or use require.js or any bundling tool:
var OneDriveAuth = require('onedrive-auth');

// create OneDrive client
var onedrive = new OneDriveAuth({
  clientId: 'YOUR-CLIENT-ID',
  scopes: 'onedrive.readonly wl.signin',
  redirectUri: 'YOUR-CALLBACK-URI',
});

// start auth process
onedrive.auth((token) => {
  // call OneDrive API endpoints with given token
});
```

## Example

Here is the simple example of authorizing using this library and interacting
with OneDrive api to create a file explorer web app:
[See it in action!](https://hlomzik.github.io/onedrive-auth)

## Files

* OneDriveExplorer (index.html) - A sample web app to view the contents of the signed in user's OneDrive and show the JSON structures returned by the API.
* OneDriveAuth (odauth.js) - A simple js library for handling the OAuth2 implicit grant flow for OneDrive. Used by the OneDriveExplorer web app.

## License

[MIT](https://github.com/hlomzik/onedrive-auth/blob/master/LICENSE)
