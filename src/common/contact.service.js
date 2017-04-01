(function() {
"use strict";

angular.module('common')
.service('ContactService', ContactService);

ContactService.$inject=['$http','ApiPath','$q', '$timeout'];
function ContactService($http,ApiPath,$q, $timeout){
  var service= this;

  var _nome='';
  var _telefone ='';
  var _email='';
  var _mensagem='';
  var _local = '';

  service.enviaEmail = function(nome, email,mensagem,telefone, local){
      var returnSave;
       _nome=nome;
       _telefone=telefone;
       _email=email;
       _mensagem=mensagem;
       _local = local;

       $http({
         url: '/enviaemail.asp',
         method: 'POST',
         data: {
           'nome': nome,
           'email': email,
           'mensagem': mensagem,
           'telefone' : telefone,
           'local': local,
         },
         headers: {'Content-Type': 'application/x-www-form-urlencoded'}
       }).
       success(function (data) {
         returnSave = true;
           //$scope.success = true;
           //exemplo de retorno: alert(data['email']);
           //$scope.user = {};
     }).
     error(function (xhr, status, error) {
       var err = eval("(" + xhr.responseText + ")");
                alert(err.Message);
           //$scope.error = true;
          returnSave = false;
       });

       return returnSave;

  };

}

})();
