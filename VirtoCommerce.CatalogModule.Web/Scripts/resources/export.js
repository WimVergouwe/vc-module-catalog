angular.module('virtoCommerce.catalogModule')
.factory('virtoCommerce.catalogModule.export', ['$resource', function ($resource) {

	return $resource('api/catalog/export/:id', { id: '@id' }, {
		run: { method: 'POST', url: 'api/catalog/export', isArray: false }
	});

}]);

angular.module('virtoCommerce.catalogModule')
.factory('virtoCommerce.catalogModule.xlsxExport', ['$resource', function($resource) {
    return $resource('api/catalog/xlsx/export/:id', { id: '@id' }, {
        run: { method: 'POST', url: 'api/catalog/xlsx/export', isArray: false}
    });
}]);