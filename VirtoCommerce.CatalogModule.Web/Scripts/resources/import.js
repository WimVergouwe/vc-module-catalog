angular.module('virtoCommerce.catalogModule')
.factory('virtoCommerce.catalogModule.import', ['$resource', function ($resource) {

	return $resource('api/catalog/import', {}, {
		getMappingConfiguration: { method: 'GET', url: 'api/catalog/import/mappingconfiguration', isArray: false },
		run: { method: 'POST', url: 'api/catalog/import', isArray: false }
	});

}]);

angular.module('virtoCommerce.catalogModule')
.factory('virtoCommerce.catalogModule.xlsxImport', ['$resource', function ($resource) {
    return $resource('api/catalog/xlsx/import/:id', { id: '@id' }, {
        getMappingConfiguration: { method: 'GET', url: 'api/catalog/xlsx/import/mappingconfiguration', isArray: false },
        run: { method: 'POST', url: 'api/catalog/xlsx/import', isArray: false }
    });
}]);