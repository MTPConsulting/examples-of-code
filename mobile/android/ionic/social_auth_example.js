(function(){
  "use strict";
  angular.module('SocialSurveyApp')
    //Factory de autenticación con google
    .factory('GooglePlugin',function($q, $window){
      return {

      /**
      * @name getDataLogin
      * @desc Login con google utilizando el plugin cordova-plugin-googleplus
      * @returns {promise}
      */
      getDataLogin: function(){
        var defered = $q.defer();
        var promise = defered.promise;

        $window.plugins.googleplus.login({
          "offline": true,
          "CLIENT_ID": "",
          "PACKAGE_NAME": ""
        },
        function (obj) {
          defered.resolve(obj);
        },
        function (msg) {
          defered.reject("Error Google: " + msg);
          }
        );

        return promise;
      },

      /**
      * @name logout
      * @desc Logout con google plus
      */
      logout: function(){
        $window.plugins.googleplus.logout(
          function (msg) {
            //console.log(msg);
          }
        );
      }
    }
    })
    //Factory de autenticación con facebook
    .factory('FacebookPlugin',function($q){
      return {

        /**
        * @name getDataLogin
        * @desc Login con google utilizando el plugin facebookConnectPlugin cordova
        * @returns {promise}
        */
        getDataLogin: function(){
          var defered = $q.defer();
          var promise = defered.promise;

          facebookConnectPlugin.login(['public_profile', 'email'], function(status) {
            facebookConnectPlugin.getAccessToken(function(token) {
              defered.resolve(token);
            }),
            function(error) {
              defered.reject("Could not get access token: " + error);
            };
          },function(error) {
            defered.reject("Error Facebook: " + error);
          });

          return promise;
        },

        /**
        * @name logout
        * @desc Login facebook
        */
        logout: function(){
          facebookConnectPlugin.logout(
            function (msg) {
              //console.log(msg);
            }
          );
        }
      }
    });
  })();
