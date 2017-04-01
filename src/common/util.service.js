(function () {
"use strict";

angular.module('common')
.service('UtilService', UtilService);


UtilService.$inject = ['$http', 'ApiPath'];
function UtilService($http, ApiPath) {
  var service = this;

  service.getUtils = function () {
    return $http.get(ApiPath + 'src/util.json').then(function (response) {
      return response.data;
    });
   };
 }

})();
