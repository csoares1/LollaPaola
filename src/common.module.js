(function() {
"use strict";

angular.module('common', [])
.constant('ApiPath', 'http://localhost:81/src/')
.config(config);

config.$inject = ['$httpProvider'];
function config($httpProvider) {
  $httpProvider.interceptors.push('loadingHttpInterceptor');
}

})();
