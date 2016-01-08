# onedrive-auth

Simple javascript OneDrive auth library. Makes an authorization and returns
back the OAuth2 token.

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

* `odauth.js` - ES2015 library, main file with simple `export default`
* `dist/odauth.js` - transpiled ES3 file to direct use in web apps; contains
  UMD block, globally accessible as `OneDriveAuth`
* `.babelrc` - options for babel@6 to get the current transpiled result

Sample OneDrive explorer page. Only available in `gh-pages` branch:

* `index.html` - a sample web app to view the contents of the signed in user's
  OneDrive and show the JSON structures returned by the API.
* `callback.html` - page assigned with `redirectUri` to work with response
  from OneDrive OAuth2 API. Returns control back to main page with token.
* `style.css`, `spinner_grey_40_transparent.gif` - just utility files for
  example page. Should be inserted in `index.html` directly later.

## License

[MIT](https://github.com/hlomzik/onedrive-auth/blob/master/LICENSE)
