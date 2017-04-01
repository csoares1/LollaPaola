(function () {
"use strict";

angular.module('public')
.controller('UtilController', UtilController);

UtilController.$inject = ['utilItems'];
function UtilController(utilItems) {
  var $ctrl = this;
  $ctrl.utilItems = utilItems;
}


})();
