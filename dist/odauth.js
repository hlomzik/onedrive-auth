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

(function (global, factory) {
  if (typeof define === "function" && define.amd) {
    define("OneDriveAuth", ["module", "exports"], factory);
  } else if (typeof exports !== "undefined") {
    factory(module, exports);
  } else {
    var mod = {
      exports: {}
    };
    factory(mod, mod.exports);
    global.OneDriveAuth = mod.exports;
  }
})(this, function (module, exports) {
  Object.defineProperty(exports, "__esModule", {
    value: true
  });

  function _classCallCheck(instance, Constructor) {
    if (!(instance instanceof Constructor)) {
      throw new TypeError("Cannot call a class as a function");
    }
  }

  var _createClass = (function () {
    function defineProperties(target, props) {
      for (var i = 0; i < props.length; i++) {
        var descriptor = props[i];
        descriptor.enumerable = descriptor.enumerable || false;
        descriptor.configurable = true;
        if ("value" in descriptor) descriptor.writable = true;
        Object.defineProperty(target, descriptor.key, descriptor);
      }
    }

    return function (Constructor, protoProps, staticProps) {
      if (protoProps) defineProperties(Constructor.prototype, protoProps);
      if (staticProps) defineProperties(Constructor, staticProps);
      return Constructor;
    };
  })();

  var OneDriveAuth = (function () {
    function OneDriveAuth(appInfo) {
      _classCallCheck(this, OneDriveAuth);

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
        appInfo.redirectOrigin = appInfo.redirectUri.match(/^[\w:]+\/\/[^\/]+/)[0];
      }

      this.appInfo = appInfo;
      var sep = this.appInfo.redirectUri.indexOf('?') < 0 ? '?' : '&';
      this.appInfo.redirectUri = this.appInfo.redirectUri.replace(/(#|$)/, sep + 'clientId=' + this.appInfo.clientId + '$1');
      this.callbacks = [];
      window.addEventListener('message', this.onAuthenticated.bind(this), false);
    }

    _createClass(OneDriveAuth, [{
      key: "auth",
      value: function auth(callback, wasClicked) {
        this.ensureHttps();
        var token = this.getTokenFromCookie();
        wasClicked = wasClicked || callback === true;
        callback = typeof callback === 'function' ? callback : false;

        if (token) {
          callback && callback(token);
          return true;
        }

        callback && this.callbacks.push(callback);

        if (wasClicked) {
          this.challengeForAuth();
        } else {
          this.showLoginButton();
        }

        return false;
      }
    }, {
      key: "ensureHttps",
      value: function ensureHttps() {
        if (window.location.protocol != "https:") {
          window.location.href = "https:" + window.location.href.substring(window.location.protocol.length);
        }
      }
    }, {
      key: "getTokenFromCookie",
      value: function getTokenFromCookie() {
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
    }, {
      key: "showLoginButton",
      value: function showLoginButton() {
        var callback = this.challengeForAuth.bind(this);
        var loginText = document.createElement('a');
        loginText.href = "#";
        loginText.id = "loginText";
        loginText.onclick = callback;
        loginText.innerText = "[sign in]";
        document.body.insertBefore(loginText, document.body.children[0]);
      }
    }, {
      key: "removeLoginButton",
      value: function removeLoginButton() {
        var loginText = document.getElementById("loginText");

        if (loginText) {
          document.body.removeChild(loginText);
        }
      }
    }, {
      key: "challengeForAuth",
      value: function challengeForAuth() {
        var appInfo = this.appInfo;
        var url = "https://login.live.com/oauth20_authorize.srf" + "?client_id=" + appInfo.clientId + "&scope=" + encodeURIComponent(appInfo.scopes) + "&response_type=token" + "&redirect_uri=" + encodeURIComponent(appInfo.redirectUri);
        this.popup(url);
      }
    }, {
      key: "popup",
      value: function popup(url) {
        var width = 525,
            height = 525,
            screenX = window.screenX,
            screenY = window.screenY,
            outerWidth = window.outerWidth,
            outerHeight = window.outerHeight;
        var left = screenX + Math.max(outerWidth - width, 0) / 2;
        var top = screenY + Math.max(outerHeight - height, 0) / 2;
        var features = ["width=" + width, "height=" + height, "top=" + top, "left=" + left, "status=no", "resizable=yes", "toolbar=no", "menubar=no", "scrollbars=yes"];
        var popup = window.open(url, "oauth", features.join(","));

        if (!popup) {
          console.error("failed to pop up auth window");
        }

        popup.focus();
      }
    }, {
      key: "onAuthenticated",
      value: function onAuthenticated(event) {
        var callback,
            token,
            data = event.data;
        if (this.appInfo.clientId !== data.clientId) return false;
        if (this.appInfo.redirectOrigin !== event.origin) return false;
        token = data.token;

        while (callback = this.callbacks.shift()) {
          callback(token);
        }
      }
    }], [{
      key: "onAuthCallback",
      value: function onAuthCallback() {
        var authInfo = this.getAuthInfoFromUrl();
        var token = authInfo["access_token"];
        var expiry = parseInt(authInfo["expires_in"]);
        var clientId = authInfo["clientId"];
        var origin = location.origin;
        this.setCookie(token, expiry);
        window.opener.postMessage({
          token: token,
          clientId: clientId
        }, origin);
        window.close();
      }
    }, {
      key: "getAuthInfoFromUrl",
      value: function getAuthInfoFromUrl() {
        if (window.location.hash) {
          var authResponse = (window.location.search + window.location.hash).substr(1);
          var authInfo = JSON.parse('{"' + authResponse.replace(/[&#]/g, '","').replace(/=/g, '":"') + '"}', function (key, value) {
            return key === "" ? value : decodeURIComponent(value);
          });
          return authInfo;
        } else {
          console.error("failed to receive auth token");
        }
      }
    }, {
      key: "setCookie",
      value: function setCookie(token, expiresInSeconds) {
        var expiration = new Date();
        expiration.setTime(expiration.getTime() + expiresInSeconds * 1000);
        var cookie = "odauth=" + token + "; path=/; expires=" + expiration.toUTCString();

        if (document.location.protocol.toLowerCase() == "https") {
          cookie = cookie + ";secure";
        }

        document.cookie = cookie;
      }
    }]);

    return OneDriveAuth;
  })();

  exports.default = OneDriveAuth;
  module.exports = exports['default'];
});