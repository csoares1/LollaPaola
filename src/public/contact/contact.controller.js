(function () {
"use strict";

angular.module('public')
.controller('contactController', contactController);

contactController.$inject = ['$scope','$http'];
function contactController($scope,$http) {
  $scope.salveSubmit = function(){

    $http({
          url: 'enviaemail.asp' ,
          method: 'POST',
          data: {
              'nome': $scope.user.nome,
              'telefone' : $scope.user.telefone,
              'email': $scope.user.email,
              'local': $scope.user.local,
              'data': $scope.user.data,
              'qtdPessoas': $scope.user.qtdPessoas,
              'mensagem': $scope.user.mensagem,
              },
        headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json'
        }
      }).
      success(function (data) {
        $scope.success = true;
        //exemplo de retorno: alert(data['email']);
        $scope.user = {};
      }).
      error(function (data) {
        console.log(data);

        $scope.error = true;

      });
  }
}

})();
