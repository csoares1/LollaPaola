(function() {
"use strict";

angular.module('common', [])
.constant('ApiPath', 'https://csoares1.github.io/LollaPaola/')
.config(config);

config.$inject = ['$httpProvider'];
function config($httpProvider) {
  $httpProvider.interceptors.push('loadingHttpInterceptor');
}

})();
