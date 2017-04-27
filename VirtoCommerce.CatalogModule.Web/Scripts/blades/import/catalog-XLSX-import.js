angular.module('virtoCommerce.catalogModule')
.controller('virtoCommerce.catalogModule.catalogExcelimportController', ['$scope', 'platformWebApp.bladeNavigationService', 'virtoCommerce.catalogModule.xlsxImport', 'platformWebApp.notifications', function ($scope, bladeNavigationService, importResourse, notificationsResource) {
	var blade = $scope.blade;
	blade.isLoading = false;
	blade.title = 'catalog.blades.catalog-Excel-import.title';

    $scope.$on("new-notification-event", function (event, notification) {
    	if (blade.notification && notification.id == blade.notification.id)
    	{
    		angular.copy(notification, blade.notification);
    	}
    });

    $scope.setForm = function (form) {
        $scope.formScope = form;
    }
  
    $scope.bladeHeadIco = 'fa fa-file-excel-o';

  
}]);
