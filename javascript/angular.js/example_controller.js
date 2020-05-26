(function () {
  "use strict";
  angular.module('SocialSurveyApp')
    .controller('HomeController', function ($filter, $ionicScrollDelegate, $scope, $timeout, FactoryGroups, FirebaseArray, getLoggedUser, NgUtilsData, NgUtilsIonic, NgUtilsSurveys, OfflineFirebase) {

      /**
  		* @name initialize
  		* @desc Inicializa los objetos y obtiene
      * las encuestas desde firebase
  		*/
      $scope.initialize = function () {
        // Inicializo los modelos a utilizar
        $scope.search = { value: "" };

        // Objeto que contendrá las encuestas
        $scope.surveys = [];
        // Para la paginación de encuestas
        $scope.limit = 10;
        $scope.init_limit = 0;

        // Para el badge de notificaciones
        $scope.badge = {
          count_notifications: 0,
          ready: false
        };

        // Para el inifitive scroll si sigo mostrando surveys
        $scope.hasMoreSurveys = true;

        // Para que no entre cuando se estan cargando las encuestas
        $scope.data_ready = false;

        // Obtengo los datos del usuario logueado
        $scope.user = getLoggedUser;
        $scope.userEncode = btoa($scope.user.email);

        // Obtengo las notificaciones
        getNotificationIsViewed();

        // Para obtener los datos en tiempo real
        sync_data();
      }

      /**
  		* @name surveys_data
  		* @desc Obtiene las encuestas, utilizando infinite scroll obteniendo de a 10 encuestas
  		*/
      $scope.surveys_data = function () {
        // Si tiene que buscar más encuestas va a firebase
        if ($scope.hasMoreSurveys) {
          // Filtro las encuestas por los publicados, que no sean grupos
          // y en un limite 10 para la paginación
          $timeout(function () {
            // Obtengo las encuestas publicas
            NgUtilsSurveys.getSurveysHomeAll($scope.user.email, $scope.surveys, $scope.init_limit, $scope.limit).then(function (obj) {
              // Si el últimi item del array no esta definido es que llego al final
              var obj_length = obj.length;
              if (typeof (obj[obj_length - 1]) == "undefined") {
                $scope.hasMoreSurveys = false;
              } else {
                $scope.surveys = obj;
                $scope.limit = $scope.limit + 10;
                $scope.init_limit = $scope.init_limit + 10;
              }

              // Indico que las encuestas estan listas
              $scope.data_ready = true;
            }, function (error) {
              $scope.$broadcast('scroll.infiniteScrollComplete');
            }).then(function (res) {
              $scope.$broadcast('scroll.infiniteScrollComplete');
            });
          }, 1000);
        } else {
          $scope.$broadcast('scroll.infiniteScrollComplete');
        }
      }

      /**
  		* @name cancel_search
  		* @desc Cancela el texto de search
  		*/
      $scope.cancel_search = function () {
        $scope.search = {
          value: ""
        };
      }

      /**
  		* @name remove_survey
  		* @desc Elimina la encuesta seleccionada
      * @param {Integer} idsurvey: Id de la encuesta
      * @param {Integer} idkeyfirebase: Id de firebase para identificar la encuesta
  		*/
      $scope.remove_survey = function (idsurvey, idkeyfirebase, idgroup) {
        var confirm = NgUtilsIonic.show_confirm("Social Survey", $filter('translate')('QUESTION_SURVEY_REMOVE'));
        confirm.then(function (res) {
          if (res) {
            // Elimino en firebase
            NgUtilsSurveys.removeSurvey(idsurvey, idkeyfirebase, idgroup);
          }
        });
      }

      /**
  		* @name getNotificationIsViewed
  		* @desc Obtengo las notificaciones pendientes para el badge
  		*/
      function getNotificationIsViewed() {
        var filters = {
          key: "is_viewed",
          value: false
        };
        NgUtilsData.getNotificationsUser(filters, $scope.user.email).then(function (data) {
          // Solo chequeo si esta online firebase y hay conexión
          OfflineFirebase.isFirebaseOnline().then(function (isOnline) {
            $scope.badge.ready = true;
            // Si hay notificaciones las cargo
            if (data.length > 0) {
              $scope.badge.count_notifications = data.length;
            } else {
              $scope.badge.count_notifications = 0;
            }
          }).catch(function (err) {
            $scope.badge.ready = false;
            $scope.badge.count_notifications = 0;
          })

        });
      }

      /**
      * @name sync_data
      * @desc Obtengo en tiempo real las nuevas encuestas o modifico y las notificaciones
      */
      function sync_data() {
        // Escucho las encuestas publicas
        var list = FirebaseArray.getRef("surveys");
        watch_data(list);

        // Escucho los grupos a los que pertenezco
        FactoryGroups.getMyIdGroupArr($scope.user.email).then(function (groups) {
          var tot = groups.length;
          for (var i = 0; i < tot; i++) {
            var list = FirebaseArray.getRef("surveys_" + groups[i]);
            watch_data(list);
          }
        });

        // Notificaciones del usuario
        var listNotif = FirebaseArray.getRef("notifications_" + btoa($scope.user.email));
        listNotif.$watch(function (event) {
          // Clave de la encuesta agregada
          var key = event.key;
          // Objeto con toda la data de la encuesta
          var snap = listNotif.$getRecord(key);

          // Chequeo que acción se realizo
          switch (event.event) {
            //Si se agrega una nueva entrada
            case "child_added":
              // Agrego una notificacion
              if ($scope.badge.ready) {
                $scope.badge.count_notifications++;
              }
              break;
          }
        });

        // Grupos, necesito la sincronizacion por si agregan un nuevo grupo
        // al que pertecence, así se visualizan las encuestas agregadas
        var listGroups = FirebaseArray.getRef("groups");
        listGroups.$watch(function (event) {
          // Clave de la encuesta agregada
          var key = event.key;
          // Objeto con toda la data de la encuesta
          var snap = listGroups.$getRecord(key);

          // Chequeo que acción se realizo
          switch (event.event) {
            // Si se agrega una nueva entrada
            case "child_added":
              // Para que no entre cuando se estan cargando las encuestas
              if ($scope.data_ready) {
                // Cargo la syncronizacion
                addSyncGroupSurvey(snap.idgroup, snap.admin);
              }
              break;
          }
        });
      }

      /**
      * @name addSyncGroupSurvey
      * @desc Agrega sincronizacion a un grupo nuevo agregado
      * @param {String} idgroup: id del grupo agregar la referencia
      * @param {String} emailAdmin: Email del admin del nuevo grupo
      */
      function addSyncGroupSurvey(idgroup, emailAdmin) {
        $timeout(function () {
          NgUtilsData.getGroupDetail({ key: "idgroup", value: idgroup }).then(function (data) {
            var contacts = data[0].contacts;
            contacts.push({ email: emailAdmin, name: "" });
            var tot = contacts.length;
            for (var i = 0; i < tot; i++) {
              // Si el grupo que se agrego pertenezco agrego la syncronizacion
              if (contacts[i].email == $scope.user.email) {
                var list = FirebaseArray.getRef("surveys_" + idgroup);
                watch_data(list);
              }
            }
          });
        }, 2000);
      }

      /**
      * @name watch_data
      * @desc Watcher para cada list
      * @param {Object} list: Coleccion a observar
      */
      function watch_data(list) {
        list.$watch(function (event) {
          // Clave de la encuesta agregada
          var key = event.key;
          // Objeto con toda la data de la encuesta
          var snap = list.$getRecord(key);

          // Chequeo que acción se realizo
          switch (event.event) {
            // Si se agrega una nueva entrada
            case "child_added":
              // Para que no entre cuando se estan cargando las encuestas
              if ($scope.data_ready) {
                //Cargo la encuesta
                $scope.surveys.unshift(snap);
              }
              break;
            // Si se modifico una entrada
            case "child_changed":
              var data = $scope.surveys.filter(function (survey) {
                return survey.$id == key
              });
              // Actualizo toda la encuesta
              try {
                data[0] = snap;
              } catch (e) { }
              break;
            // Elimino en la vista en tiempo real el objeto eliminado
            case "child_removed":
              // Busco el objecto a eliminar y lo elimino
              $scope.surveys.filter(function (survey, index) {
                if (survey.$id == key) {
                  $scope.surveys.splice(index, 1);
                }
              });
              break;
          }
        });
      }

      // Evento para detectar los cambio de las urls.
      $scope.$on('$stateChangeSuccess', function (ev, to, toParams, from, fromParams) {
        var url = from.name;
        switch (url) {
          // Para saber si viene desde los filtros o el perfil
          // Para saber si reinicio las encuestas
          case "app.profile":
            // Variables de inicialización
            $scope.initialize();
            // Mueve el scroll para que ejecute el loading y cargue las encuestas
            $ionicScrollDelegate.scrollTop();
            break;
          case "app.seesurveysfilters":
            // Variables de inicialización
            $scope.initialize();
            // Mueve el scroll para que ejecute el loading y cargue las encuestas
            $ionicScrollDelegate.scrollTop();
            break;
          // Si viene desde notificaciones inicializo el badge.count
          case "app.notifications":
            $scope.badge.count_notifications = 0;
            break;
        }
      });
    });
})();
