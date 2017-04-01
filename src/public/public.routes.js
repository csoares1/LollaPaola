(function() {
'use strict';

angular.module('public')
.config(routeConfig);

/**
 * Configures the routes and views
 */
routeConfig.$inject = ['$stateProvider'];
function routeConfig ($stateProvider) {
  // Routes
  $stateProvider
    .state('public', {
      absract: true,
      templateUrl: 'src/public/public.html'
    })
    .state('public.home', {
      url: '/',
      templateUrl: 'src/public/home/home.html'
    })
    .state('public.menu', {
      url: '/menu',
      templateUrl: 'src/public/menu/menu.html',
      controller: 'MenuController',
      controllerAs: 'menuCtrl',
      resolve: {
        menuCategories: ['MenuService', function (MenuService) {
          return MenuService.getCategories();
        }]
      }
    })
    .state('public.util', {
      url: '/util',
      templateUrl: 'src/public/util/util.html',
      controller: 'UtilController',
      controllerAs: 'utilCtrl',
      resolve:{
         utilItems:['UtilService', function(UtilService){
           return UtilService.getUtils();
         }]
        }
    })
    .state('public.menuitems', {
      url: '/menu/{category}',
      templateUrl: 'src/public/menu-items/menu-items.html',
      controller: 'MenuItemsController',
      controllerAs: 'menuItemsCtrl',
      resolve: {
        menuItems: ['$stateParams','MenuService', function ($stateParams, MenuService) {
          return MenuService.getMenuItems($stateParams.category);
        }]
      }
    })
    .state('public.contact',{
      url: '/contact',
      templateUrl:'src/public/contact/contact.html',
      controller:'contactController',
      controllerAs:'contactCtrl'
    })

    .state('public.myinfo',{
      url: '/myinfo',
      templateUrl:'src/public/my-info/my-info.html'
    });


}
})();
