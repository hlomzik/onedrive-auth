"use strict";

/**
 * Small lib to go through the OneDrive authentication process and get back the token.
 * 
 * Instructions:
 * - Host a copy of callback.html and odauth.js on your domain.
 * - Embed odauth.js in your app like this:
 *   <script src="odauth.js"></script>
 * - Define the onAuthenticated(token) function in your app to receive the auth token.
 * - Create instance of OneDriveAuth and call auth() method to begin and whenever you
 *   need an auth token to make an API call. If you're making an api call in response to
 *   a user's click action, call auth(true), otherwise just call auth(). The difference
 *   is that sometimes OneDriveAuth needs to pop up a window so the user can sign in,
 *   grant your app permission, etc. The pop up can only be launched in response
 *   to a user click, otherwise the browser's popup blocker will block it. When
 *   OneDriveAuth isn't called in click mode, it'll put a sign-in button at the top of
 *   your page for the user to click. When it's done, it'll remove that button.
 */
class OneDriveAuth {
  /**
   * @constructor
   * @param {Array} appInfo
   * @param {string} appInfo.clientId from App Settings
   * @param {string} appInfo.scopes separated by space ("onedrive.readonly wl.signin" for example)
   * @param {string} appInfo.redirectUri
   */
  constructor(appInfo) {
    if (!appInfo.clientId) {
      throw "appInfo object should have `clientId` property set to your application id";
    }
    
    if (!appInfo.scopes) {
      throw "appInfo object should have `scopes` property set to the scopes your app needs";
    }
    
    if (!appInfo.redirectUri) {
      throw "appInfo object should have `redirectUri` property set to your redirect landing url";
    }
    this.appInfo = appInfo;
  }
  
  /**
   * The main auth method. First, we check if we have the user's auth token stored
   * in a cookie. if so, we read it and immediately call your onAuthenticated() method.
   * If we can't find the auth cookie, we need to pop up a window and send the user
   * to Microsoft Account so that they can sign in or grant your app the permissions
   * it needs. Depending on whether or not auth() was called in response to a user
   * click, it will either pop up the auth window or display a sign-in button for
   * the user to click (which results in the pop-up). when the user finishes the
   * auth flow, the popup window redirects back to your hosted callback.html file,
   * which calls the onAuthCallback() static method below. It then sets the auth cookie
   * and calls your app's onAuthenticated() function, passing in the optional 'window'
   * argument for the popup window. your onAuthenticated function should close the
   * popup window if it's passed in.
   *
   * Subsequent calls to auth() will usually complete immediately without the
   * popup because the cookie is still fresh.
   * 
   * @param {boolean} wasClicked whether the call is result of a click or not
   */
  auth(wasClicked) {
    ensureHttps();
    var token = this.getTokenFromCookie();
    if (token) {
      window.onAuthenticated(token);
    } else if (wasClicked) {
      this.challengeForAuth();
    } else {
      this.showLoginButton();
    }
  }
  
  /**
   * For added security we require https
   */
  ensureHttps() {
    if (window.location.protocol != "https:") {
      window.location.href = "https:" + window.location.href.substring(window.location.protocol.length);
    }
  }
  
  getTokenFromCookie() {
    var cookies = document.cookie;
    var name = "odauth=";
    var start = cookies.indexOf(name);
    if (start >= 0) {
      start += name.length;
      var end = cookies.indexOf(';', start);
      if (end < 0) {
        end = cookies.length;
      }
      
      var value = cookies.substring(start, end);
      return value;
    }
    
    return "";
  }
  
  /**
   * Called when a login button needs to be displayed for the user to click on.
   * If a customLoginButton() function is defined by your app, it will be called
   * with `callback` passed in to indicate the button should be added. Otherwise, it
   * will insert a textual login link at the top of the page. If defined, your
   * showCustomLoginButton should call `callback` when clicked.
   */
  showLoginButton() {
    var callback = this.challengeForAuth.bind(this);
    if (typeof window.showCustomLoginButton === "function") {
      window.showCustomLoginButton(callback);
      return;
    }
    
    var loginText = document.createElement('a');
    loginText.href = "#";
    loginText.id = "loginText";
    loginText.onclick = callback;
    loginText.innerText = "[sign in]";
    document.body.insertBefore(loginText, document.body.children[0]);
  }
  
  /**
   * Called with the login button created by showLoginButton() needs to be
   * removed. If a customLoginButton() function is defined by your app, it will
   * be called with 'false' passed in to indicate the button should be removed.
   * Otherwise it will remove the textual link that showLoginButton() created.
   */
  removeLoginButton() {
    if (typeof window.showCustomLoginButton === "function") {
      window.showCustomLoginButton(false);
      return;
    }
    
    var loginText = document.getElementById("loginText");
    if (loginText) {
      document.body.removeChild(loginText);
    }
  }
  
  /**
   * Send user to the Microsoft Account OAuth2.0 page to authorize the app
   * and return the user to your redirectUri.
   */
  challengeForAuth() {
    var appInfo = this.appInfo;
    var url =
      "https://login.live.com/oauth20_authorize.srf" +
      "?client_id=" + appInfo.clientId +
      "&scope=" + encodeURIComponent(appInfo.scopes) +
      "&response_type=token" +
      "&redirect_uri=" + encodeURIComponent(appInfo.redirectUri);
    this.popup(url);
  }
  
  popup(url) {
    var width = 525,
      height = 525,
      screenX = window.screenX,
      screenY = window.screenY,
      outerWidth = window.outerWidth,
      outerHeight = window.outerHeight;
    
    var left = screenX + Math.max(outerWidth - width, 0) / 2;
    var top = screenY + Math.max(outerHeight - height, 0) / 2;
    
    var features = [
      "width=" + width,
      "height=" + height,
      "top=" + top,
      "left=" + left,
      "status=no",
      "resizable=yes",
      "toolbar=no",
      "menubar=no",
      "scrollbars=yes"
    ];
    var popup = window.open(url, "oauth", features.join(","));
    if (!popup) {
      console.error("failed to pop up auth window");
    }
    
    popup.focus();
  }
  
  /**
   * Called from the callback page after OAuth authorization.
   * On success it save the token to cookie and call user's onAuthenticated()
   * function in parent window.
   */
  static onAuthCallback() {
    var authInfo = this.getAuthInfoFromUrl();
    var token = authInfo["access_token"];
    var expiry = parseInt(authInfo["expires_in"]);
    this.setCookie(token, expiry);
    window.opener.onAuthenticated(token, window);
  }
  
  static getAuthInfoFromUrl() {
    if (window.location.hash) {
      var authResponse = window.location.hash.substring(1);
      var authInfo = JSON.parse(
        '{"' + authResponse.replace(/&/g, '","').replace(/=/g, '":"') + '"}',
        (key, value) => key === "" ? value : decodeURIComponent(value)
      );
      return authInfo;
    } else {
      console.error("failed to receive auth token");
    }
  }
  
  static setCookie(token, expiresInSeconds) {
    var expiration = new Date();
    expiration.setTime(expiration.getTime() + expiresInSeconds * 1000);
    var cookie = "odauth=" + token + "; path=/; expires=" + expiration.toUTCString();
    
    if (document.location.protocol.toLowerCase() == "https") {
      cookie = cookie + ";secure";
    }
    
    document.cookie = cookie;
  }
}
