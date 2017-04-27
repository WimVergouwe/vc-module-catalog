using System;
using System.Collections.Generic;
using System.IO;
using CsvHelper;
using VirtoCommerce.CatalogModule.Data.Repositories;
using VirtoCommerce.Domain.Catalog.Model;
using VirtoCommerce.Domain.Catalog.Services;
using VirtoCommerce.Domain.Commerce.Services;
using VirtoCommerce.Domain.Inventory.Services;
using VirtoCommerce.Domain.Pricing.Services;
using VirtoCommerce.Platform.Core.ExportImport;

namespace VirtoCommerce.CatalogModule.Web.ExportImport.Csv
{
    public sealed class CsvCatalogImporter : AbstractCatalogImporter
    {
        private readonly ICatalogService _catalogService;

        public CsvCatalogImporter(ICatalogService catalogService, ICategoryService categoryService, IItemService productService,
                                  ISkuGenerator skuGenerator,
                                  IPricingService pricingService, IInventoryService inventoryService, ICommerceService commerceService,
                                  IPropertyService propertyService, ICatalogSearchService searchService, Func<ICatalogRepository> catalogRepositoryFactory) : base(categoryService, productService, skuGenerator, pricingService, inventoryService, commerceService, propertyService, searchService, catalogRepositoryFactory)
        {
            _catalogService = catalogService;
        }

        public override void DoImport(Stream inputStream, ImportInfo importInfo, Action<ExportImportProgressInfo> progressCallback)
        {
            var products = new List<CatalogProduct>();

            var progressInfo = new ExportImportProgressInfo
            {
                Description = "Reading products from csv..."
            };
            progressCallback(progressInfo);

            using (var reader = new CsvReader(new StreamReader(inputStream)))
            {
                var configuration = (importInfo as CsvImportInfo)?.Configuration;

                reader.Configuration.Delimiter = configuration?.Delimiter ?? ",";
                reader.Configuration.RegisterClassMap(new CsvProductMap(configuration));
                reader.Configuration.WillThrowOnMissingField = false;

                while (reader.Read())
                {
                    try
                    {
                        var product = reader.GetRecord<CsvProduct>();
                        products.Add(product);
                    }
                    catch (Exception ex)
                    {
                        var error = ex.Message;
                        if (ex.Data.Contains("CsvHelper"))
                        {
                            error += ex.Data["CsvHelper"];
                        }
                        progressInfo.Errors.Add(error);
                        progressCallback(progressInfo);
                    }
                }
            }

            var catalog = _catalogService.GetById(importInfo.CatalogId);

            SaveCategoryTree(catalog, products, progressInfo, progressCallback);
            SaveProducts(catalog, products, progressInfo, progressCallback);
        }
    }
}
