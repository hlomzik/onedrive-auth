"use strict";

/**
 * Small lib to go through the OneDrive authentication process and get back the token.
 * 
 * Instructions:
 * - Host a copy of callback.html and odauth.js on your domain.
 * - Embed odauth.js in your app like this:
 *   <script src="odauth.js"></script>
 * - Create instance of OneDriveAuth and call auth() method to begin and whenever you
 *   need to make an API call. In first `callback` param you should provide handler of
 *   successful authorization. If you're making an api call in response to a user's click
 *   action, call auth(callback, true), otherwise just call auth(callback). The difference
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
   * @param {string} appInfo.redirectUri for callback popup
   * @param {string} [appInfo.redirectOrigin] origin of callback window
   * @param {boolean} [appInfo.requireHttps=true] check for https and throw an error if it's omitted
   */
  constructor(appInfo) {
    this.appInfo = Object.assign({}, appInfo); // clone parameters

    if (!appInfo.clientId) {
      throw "appInfo object should have `clientId` property set to your application id";
    }
    
    if (!appInfo.scopes) {
      throw "appInfo object should have `scopes` property set to the scopes your app needs";
    }
    
    if (!appInfo.redirectUri) {
      throw "appInfo object should have `redirectUri` property set to your redirect landing url";
    }
    
    if (!appInfo.redirectOrigin) {
      // get the scheme://host:port from redirectUri
      this.appInfo.redirectOrigin = appInfo.redirectUri.match(/^[\w:]+\/\/[^\/]+/)[0];
    }
    
    if (typeof appInfo.requireHttps === 'undefined') {
      this.appInfo.requireHttps = true;
    }
    
    var sep = this.appInfo.redirectUri.indexOf('?') < 0 ? '?' : '&';
    // adds ?clientId=... to the end of string or before the hash in the redirectUri
    // because the auth window doesn't get back any info about client from OneDrive
    this.appInfo.redirectUri = this.appInfo.redirectUri.replace(/(#|$)/, sep + 'clientId=' + this.appInfo.clientId + '$1');
    // list of handlers for successful authorization
    this.callbacks = [];
  }
  
  /**
   * Callback to handle successful authorization
   * @callback onAuth
   * @param {string} token
   */
  
  /**
   * The main auth method. First, we check if we have the user's auth token stored
   * in a cookie. if so, we read it and immediately call your `callback` method.
   * If we can't find the auth cookie, we need to pop up a window and send the user
   * to Microsoft Account so that they can sign in or grant your app the permissions
   * it needs. Depending on whether or not auth() was called in response to a user
   * click, it will either pop up the auth window or display a sign-in button for
   * the user to click (which results in the pop-up). when the user finishes the
   * auth flow, the popup window redirects back to your hosted callback.html file,
   * which calls the onAuthCallback() static method below. It then sets the auth cookie
   * and calls given `callback` function after closing the auth window.
   *
   * Subsequent calls to auth() will usually complete immediately without the
   * popup because the cookie is still fresh.
   * 
   * @param {onAuth} [callback] to call in case of success authorization; if missed promise is returned
   * @param {boolean} [wasClicked=false] whether the call is result of a click or not
   * @return {boolean|Promise.<string>}
   *  `true` if authorized yet;
   *  `false` if authorization started or in progress;
   *  `Promise` if `callback` is missed
   * @throws if https is required but omitted
   */
  auth(callback, wasClicked) {
    this.ensureHttps();
    var token = this.getTokenFromCookie();
    // deal with both optional params
    wasClicked = wasClicked || (callback === true);
    callback = (typeof callback === 'function') ? callback : null;
    
    if (token) {
      if (callback) {
        callback(token);
        return true;
      } else {
        return Promise.resolve(token);
      }
    }
    
    // would be called in OneDriveAuth.onAuthenticated() method
    callback && this.callbacks.push(callback);
    
    if (this.inProgress) return false;
    this.inProgress = true;
    
    if (wasClicked) {
      this.challengeForAuth();
      
      return new Promise((ok, no) => {
        // listen for callback's message with auth token
        window.addEventListener('message', e => {
          let p = this.onAuthenticated(e);
          p && p.then(ok, no);
        }, false);
      });
    } else if (!callback) {
      return Promise.reject();
    }
    return false;
  }
  
  /**
   * Check if current page loaded via HTTPS
   * @return {boolean}
   */
  static isHttps() {
    return window.location.protocol.toLowerCase() === "https:";
  }
  
  /**
   * For added security we check for https
   * @throws if https is required but omitted
   */
  ensureHttps() {
    if (this.appInfo.requireHttps && !this.isHttps()) {
      throw new Error("HTTPS is required to authorize this application for OneDrive");
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
   * Called from auth window in response to successful authorization
   * @param {Event} event object with sent data
   * @param {string} event.data.token
   * @param {string} event.data.clientId
   * @param {string} event.origin origin of callback popup window
   * @return {boolean|Promise.<string, Error>} token on success
   */
  onAuthenticated(event) {
    var callback, data = event.data, token = data.token;
    // skip the message addressed to another client or came from unknown location
    if (this.appInfo.clientId !== data.clientId) return false;
    if (this.appInfo.redirectOrigin !== event.origin) return false;
    
    this.inProgress = false;
    
    if (!token) {
      let error = new Error();
      return Promise.reject(error);
    } else {
      while (callback = this.callbacks.shift()) {
        callback(token);
      }
      return Promise.resolve(token);
    }
  }
  
  /**
   * Called from the callback page after OAuth authorization.
   * On success it save the token to cookie and call onAuthenticated() method
   * of corresponding OneDriveAuth instance in parent window.
   */
  static onAuthCallback() {
    var authInfo = this.getAuthInfoFromUrl();
    var token = authInfo["access_token"];
    var expiry = parseInt(authInfo["expires_in"]);
    var clientId = authInfo["clientId"];
    var origin = location.origin;
    
    this.setCookie(token, expiry);
    window.opener.postMessage({ token, clientId }, origin);
    window.close();
  }
  
  static getAuthInfoFromUrl() {
    if (window.location.hash) {
      var authResponse = (window.location.search + window.location.hash).substr(1);
      var authInfo = JSON.parse(
        '{"' + authResponse.replace(/[&#]/g, '","').replace(/=/g, '":"') + '"}',
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
    
    if (this.isHttps()) {
      cookie = cookie + ";secure";
    }
    
    document.cookie = cookie;
  }
}

export default OneDriveAuth;
