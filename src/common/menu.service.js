(function () {
"use strict";

angular.module('common')
.service('MenuService', MenuService);


MenuService.$inject = ['$http', 'ApiPath'];
function MenuService($http, ApiPath) {
  var service = this;

  service.getCategories = function () {
    return $http.get(ApiPath + 'src/categories.json').then(function (response) {
      return response.data;
    });
  };

  service.getUtils = function () {
    return $http.get(ApiPath + 'src/util.json').then(function (response) {
      return response.data;
    });
  };


  service.getMenuItems = function (category) {
  //  var config = {};
  //  if (category) {
  //    config.params = {'category': category};
  //  }

    return $http.get(ApiPath +'src/'+ category +'.json').then(function (response) {
      return response.data;
    });
  };

}



})();
