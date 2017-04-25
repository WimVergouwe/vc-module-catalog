﻿//Call this to register our module to main application
var catalogsModuleName = "virtoCommerce.catalogModule";

if (AppDependencies != undefined) {
    AppDependencies.push(catalogsModuleName);
}

angular.module(catalogsModuleName, [
])
.config(
  ['$stateProvider', function ($stateProvider) {
      $stateProvider
          .state('workspace.catalog', {
              url: '/catalog',
              templateUrl: '$(Platform)/Scripts/common/templates/home.tpl.html',
              controller: [
                  '$scope', 'platformWebApp.bladeNavigationService', function ($scope, bladeNavigationService) {

                      var blade = {
                          id: 'categories',
                          title: 'catalog.blades.catalogs-list.title',
                          breadcrumbs: [],
                          subtitle: 'catalog.blades.catalogs-list.subtitle',
                          controller: 'virtoCommerce.catalogModule.catalogsListController',
                          template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/catalogs-list.tpl.html',
                          isClosingDisabled: true
                      };
                      bladeNavigationService.showBlade(blade);
                      $scope.moduleName = 'vc-catalog';
                  }
              ]
          });
  }
  ]
)
// define search filters to be accessible platform-wide
.factory('virtoCommerce.catalogModule.predefinedSearchFilters', ['$localStorage', function ($localStorage) {
    $localStorage.catalogSearchFilters = $localStorage.catalogSearchFilters || [];

    return {
        register: function (currentFiltersUpdateTime, currentFiltersStorageKey, newFilters) {
            _.each(newFilters, function (newFilter) {
                var found = _.find($localStorage.catalogSearchFilters, function (x) {
                    return x.id == newFilter.id;
                });
                if (found) {
                    if (found && (!found.lastUpdateTime || found.lastUpdateTime < currentFiltersUpdateTime)) {
                        angular.copy(newFilter, found);
                    }
                } else if (!$localStorage[currentFiltersStorageKey] || $localStorage[currentFiltersStorageKey] < currentFiltersUpdateTime) {
                    $localStorage.catalogSearchFilters.splice(0, 0, newFilter);
                }
            });

            $localStorage[currentFiltersStorageKey] = currentFiltersUpdateTime;
        }
    };
}])
.run(
  ['platformWebApp.authService', 'platformWebApp.mainMenuService', 'platformWebApp.widgetService', '$state', 'platformWebApp.pushNotificationTemplateResolver', 'platformWebApp.bladeNavigationService', 'virtoCommerce.catalogModule.catalogImportService', 'virtoCommerce.catalogModule.catalogExportService', 'platformWebApp.permissionScopeResolver', 'virtoCommerce.catalogModule.catalogs', 'virtoCommerce.catalogModule.predefinedSearchFilters',
	function (authService, mainMenuService, widgetService, $state, pushNotificationTemplateResolver, bladeNavigationService, catalogImportService, catalogExportService, scopeResolver, catalogs, predefinedSearchFilters) {

	    //Register module in main menu
	    var menuItem = {
	        path: 'browse/catalog',
	        icon: 'fa fa-folder',
	        title: 'catalog.main-menu-title',
	        priority: 20,
	        action: function () { $state.go('workspace.catalog'); },
	        permission: 'catalog:access'
	    };
	    mainMenuService.addMenuItem(menuItem);

	    //NOTIFICATIONS
	    // CsvExport
	    var menuCsvExportTemplate =
            {
                priority: 900,
                satisfy: function (notify, place) { return place == 'menu' && notify.notifyType == 'CatalogCsvExport'; },
                template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/export/notifications/menuExport.tpl.html',
                action: function (notify) { $state.go('workspace.pushNotificationsHistory', notify) }
            };
	    pushNotificationTemplateResolver.register(menuCsvExportTemplate);

	    var historyCsvExportTemplate =
          {
              priority: 900,
              satisfy: function (notify, place) { return place == 'history' && notify.notifyType == 'CatalogCsvExport'; },
              template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/export/notifications/historyExport.tpl.html',
              action: function (notify) {
                  var blade = {
                      id: 'CatalogCsvExportDetail',
                      title: 'catalog export detail',
                      subtitle: 'detail',
                      template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/export/catalog-CSV-export.tpl.html',
                      controller: 'virtoCommerce.catalogModule.catalogCSVexportController',
                      notification: notify
                  };
                  bladeNavigationService.showBlade(blade);
              }
          };
	    pushNotificationTemplateResolver.register(historyCsvExportTemplate);
	    //Import
	    var menuCsvImportTemplate =
          {
              priority: 900,
              satisfy: function (notify, place) { return place == 'menu' && notify.notifyType == 'CatalogCsvImport'; },
              template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/import/notifications/menuImport.tpl.html',
              action: function (notify) { $state.go('workspace.pushNotificationsHistory', notify) }
          };
	    pushNotificationTemplateResolver.register(menuCsvImportTemplate);

	    var historyCsvImportTemplate =
          {
              priority: 900,
              satisfy: function (notify, place) { return place == 'history' && notify.notifyType == 'CatalogCsvImport'; },
              template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/import/notifications/historyImport.tpl.html',
              action: function (notify) {
                  var blade = {
                      id: 'CatalogCsvImportDetail',
                      title: 'catalog import detail',
                      subtitle: 'detail',
                      template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/import/catalog-CSV-import.tpl.html',
                      controller: 'virtoCommerce.catalogModule.catalogCSVimportController',
                      notification: notify
                  };
                  bladeNavigationService.showBlade(blade);
              }
          };
	    pushNotificationTemplateResolver.register(historyCsvImportTemplate);

	    //Register dashboard widgets
	    //widgetService.registerWidget({
	    //    isVisible: function () { return authService.checkPermission('catalog:???'); },
	    //    controller: 'virtoCommerce.catalogModule.dashboard.catalogsWidgetController',
	    //    template: 'tile-count.html'
	    //}, 'mainDashboard');
	    //widgetService.registerWidget({
	    //    isVisible: function () { return authService.checkPermission('catalog:???'); },
	    //    controller: 'virtoCommerce.catalogModule.dashboard.productsWidgetController',
	    //    template: 'tile-count.html'
	    //}, 'mainDashboard');

	    //Register image widget
	    var itemImageWidget = {
	        controller: 'virtoCommerce.catalogModule.itemImageWidgetController',
	        size: [2, 2],
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/itemImageWidget.tpl.html'
	    };
	    widgetService.registerWidget(itemImageWidget, 'itemDetail');
	    //Register item property widget
	    var itemPropertyWidget = {
	        controller: 'virtoCommerce.catalogModule.itemPropertyWidgetController',
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/itemPropertyWidget.tpl.html'
	    };
	    widgetService.registerWidget(itemPropertyWidget, 'itemDetail');

	    //Register item associations widget
	    var itemAssociationsWidget = {
	        controller: 'virtoCommerce.catalogModule.itemAssociationsWidgetController',
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/itemAssociationsWidget.tpl.html'
	    };
	    widgetService.registerWidget(itemAssociationsWidget, 'itemDetail');

	    //Register item seo widget
	    var itemSeoWidget = {
	        controller: 'virtoCommerce.coreModule.seo.seoWidgetController',
	        template: 'Modules/$(VirtoCommerce.Core)/Scripts/SEO/widgets/seoWidget.tpl.html',
	        objectType: 'CatalogProduct',
	        getDefaultContainerId: function (blade) { return undefined; },
	        getLanguages: function (blade) { return _.pluck(blade.catalog.languages, 'languageCode'); }
	    };
	    widgetService.registerWidget(itemSeoWidget, 'itemDetail');

	    //seoObjectBladesResolver.registerBladeForSeoObjectType('catalogProduct', function (seoInfo) {
	    //	return {
	    //		id: "product-detail",
	    //		itemId: seoInfo.objectId,
	    //		controller: 'virtoCommerce.catalogModule.itemDetailController',
	    //		template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/item-detail.tpl.html'
	    //	};
	    //});
	    //seoObjectBladesResolver.registerBladeForSeoObjectType('category', function (seoInfo) {
	    //	return {
	    //		id: "listCategoryDetail",
	    //		currentEntityId: seoInfo.objectId,
	    //		controller: 'virtoCommerce.catalogModule.categoryDetailController',
	    //		template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/category-detail.tpl.html',
	    //	};
	    //});

	    //Register dimensions widget
	    var dimensionsWidget = {
	        controller: 'virtoCommerce.catalogModule.itemDimensionWidgetController',
	        isVisible: function (blade) { return blade.productType == 'Physical'; },
	        size: [2, 1],
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/itemDimensionWidget.tpl.html'
	    };
	    widgetService.registerWidget(dimensionsWidget, 'itemDetail');
	    //Register item editorialReview widget
	    var editorialReviewWidget = {
	        controller: 'virtoCommerce.catalogModule.editorialReviewWidgetController',
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/editorialReviewWidget.tpl.html'
	    };
	    widgetService.registerWidget(editorialReviewWidget, 'itemDetail');

	    //Register variation widget
	    var variationWidget = {
	        controller: 'virtoCommerce.catalogModule.itemVariationWidgetController',
	        isVisible: function (blade) { return blade.id !== 'variationDetail'; },
	        size: [1, 1],
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/itemVariationWidget.tpl.html'
	    };
	    widgetService.registerWidget(variationWidget, 'itemDetail');
	    //Register asset widget
	    var itemAssetWidget = {
	        controller: 'virtoCommerce.catalogModule.itemAssetWidgetController',
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/itemAssetWidget.tpl.html'
	    };
	    widgetService.registerWidget(itemAssetWidget, 'itemDetail');

	    //Register widgets to categoryDetail
	    widgetService.registerWidget(itemImageWidget, 'categoryDetail');

	    var categoryPropertyWidget = {
	        controller: 'virtoCommerce.catalogModule.categoryPropertyWidgetController',
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/categoryPropertyWidget.tpl.html'
	    };
	    widgetService.registerWidget(categoryPropertyWidget, 'categoryDetail');

	    //Register category seo widget
	    var categorySeoWidget = {
	        controller: 'virtoCommerce.coreModule.seo.seoWidgetController',
	        template: 'Modules/$(VirtoCommerce.Core)/Scripts/SEO/widgets/seoWidget.tpl.html',
	        objectType: 'Category',
	        getDefaultContainerId: function (blade) { return undefined; },
	        getLanguages: function (blade) { return _.pluck(blade.catalog.languages, 'languageCode'); }
	    };
	    widgetService.registerWidget(categorySeoWidget, 'categoryDetail');

	    //Register catalog widgets
	    var catalogLanguagesWidget = {
	        controller: 'virtoCommerce.catalogModule.catalogLanguagesWidgetController',
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/catalogLanguagesWidget.tpl.html'
	    };
	    widgetService.registerWidget(catalogLanguagesWidget, 'catalogDetail');

	    var catalogPropertyWidget = {
	        isVisible: function (blade) { return !blade.isNew; },
	        controller: 'virtoCommerce.catalogModule.catalogPropertyWidgetController',
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/widgets/catalogPropertyWidget.tpl.html'
	    };
	    widgetService.registerWidget(catalogPropertyWidget, 'catalogDetail');

	    // IMPORT / EXPORT
	    catalogImportService.register({
	        name: 'VirtoCommerce CSV import',
	        description: 'Native VirtoCommerce catalog data import from CSV',
	        icon: 'fa fa-file-archive-o',
	        controller: 'virtoCommerce.catalogModule.catalogCSVimportWizardController',
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/import/wizard/catalog-CSV-import-wizard.tpl.html'
	    });
	    catalogExportService.register({
	        name: 'VirtoCommerce CSV export',
	        description: 'Native VirtoCommerce catalog data export to CSV',
	        icon: 'fa fa-file-archive-o',
	        controller: 'virtoCommerce.catalogModule.catalogCSVexportController',
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/export/catalog-CSV-export.tpl.html'
	    });
	    catalogExportService.register({
	        name: 'Excel export',
	        description: 'Catalog data export to Excel',
	        icon: 'fa fa-file-excel-o',
	        controller: 'virtoCommerce.catalogModule.catalogExcelexportController', //extension blade controller which will be displayed when user  choose you extension 
	        template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/export/catalog-XLSX-export.tpl.html' //template for extension blade
	    });

	    //Security scopes
	    //Register permission scopes templates used for scope bounded definition in role management ui

	    var catalogSelectScope = {
	        type: 'CatalogSelectedScope',
	        title: 'catalog.permissions.catalog-scope.title',
	        selectFn: function (blade, callback) {
	            var newBlade = {
	                id: 'catalog-pick',
	                title: this.title,
	                subtitle: 'catalog.permissions.catalog-scope.blade.subtitle',
	                currentEntity: this,
	                onChangesConfirmedFn: callback,
	                dataPromise: catalogs.query().$promise,
	                controller: 'platformWebApp.security.scopeValuePickFromSimpleListController',
	                template: '$(Platform)/Scripts/app/security/blades/common/scope-value-pick-from-simple-list.tpl.html'
	            };
	            bladeNavigationService.showBlade(newBlade, blade);
	        }
	    };
	    scopeResolver.register(catalogSelectScope);

	    var categorySelectScope = {
	        type: 'CatalogSelectedCategoryScope',
	        title: 'catalog.permissions.category-scope.title',
	        selectFn: function (blade, callback) {
	            var selectedListItems = _.map(this.assignedScopes, function (x) { return { id: x.scope, name: x.label }; });
	            var options = {
	                showCheckingMultiple: false,
	                allowCheckingItem: false,
	                allowCheckingCategory: true,
	                selectedItemIds: _.map(this.assignedScopes, function (x) { return x.scope; }),
	                checkItemFn: function (listItem, isSelected) {
	                    if (isSelected) {
	                        if (_.all(selectedListItems, function (x) { return x.id != listItem.id; })) {
	                            selectedListItems.push(listItem);
	                        }
	                    }
	                    else {
	                        selectedListItems = _.reject(selectedListItems, function (x) { return x.id == listItem.id; });
	                    }
	                }
	            };
	            var scopeOriginal = this.scopeOriginal;
	            var newBlade = {
	                id: "CatalogItemsSelect",
	                title: "catalog.blades.catalog-items-select.title",
	                controller: 'virtoCommerce.catalogModule.catalogItemSelectController',
	                template: 'Modules/$(VirtoCommerce.Catalog)/Scripts/blades/common/catalog-items-select.tpl.html',
	                options: options,
	                breadcrumbs: [],
	                toolbarCommands: [
                      {
                          name: "platform.commands.confirm",
                          icon: 'fa fa-plus',
                          executeMethod: function (blade) {
                              var scopes = _.map(selectedListItems, function (x) {
                                  return angular.extend({ scope: x.id, label: x.name }, scopeOriginal);
                              });
                              callback(scopes);
                              bladeNavigationService.closeBlade(blade);

                          },
                          canExecuteMethod: function () {
                              return selectedListItems.length > 0;
                          }
                      }]
	            };
	            bladeNavigationService.showBlade(newBlade, blade);
	        }
	    };
	    scopeResolver.register(categorySelectScope);

	    // predefine search filters for catalog search
	    predefinedSearchFilters.register(1477584000000, 'catalogSearchFiltersDate', [{ name: 'catalog.blades.categories-items-list.labels.filter-new' }]);
	}]);
