using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CsvHelper;
using VirtoCommerce.CatalogModule.Data.Repositories;
using VirtoCommerce.CatalogModule.Web.Utilities;
using VirtoCommerce.Domain.Catalog.Model;
using VirtoCommerce.Domain.Catalog.Services;
using VirtoCommerce.Domain.Commerce.Services;
using VirtoCommerce.Domain.Inventory.Services;
using VirtoCommerce.Domain.Pricing.Services;
using VirtoCommerce.Platform.Core.ExportImport;

namespace VirtoCommerce.CatalogModule.Web.ExportImport.Xlsx
{
    public class XlsxCatalogImporter : AbstractCatalogImporter
    {
        private readonly ICatalogService _catalogService;

        public XlsxCatalogImporter(ICatalogService catalogService, ICategoryService categoryService, IItemService productService,
                                  ISkuGenerator skuGenerator,
                                  IPricingService pricingService, IInventoryService inventoryService, ICommerceService commerceService,
                                  IPropertyService propertyService, ICatalogSearchService searchService, Func<ICatalogRepository> catalogRepositoryFactory) : base(categoryService, productService, skuGenerator, pricingService, inventoryService, commerceService, propertyService, searchService, catalogRepositoryFactory)
        {
            _catalogService = catalogService;
        }

        public override void DoImport(Stream inputStream, ImportInfo importInfo, Action<ExportImportProgressInfo> progressCallback)
        {
            var progressInfo = new ExportImportProgressInfo
            {
                Description = "Reading products from Excel file..."
            };
            progressCallback(progressInfo);
            
            var products = DemoXlsxUtilities.GetProductsFromFile(inputStream, GetImportDefinition()).ToList();

            var catalog = _catalogService.GetById(importInfo.CatalogId);

            SaveCategoryTree(catalog, products, progressInfo, progressCallback);
            SaveProducts(catalog, products, progressInfo, progressCallback);
        }

        private ImportDefinition<XlsxProduct> GetImportDefinition()
        {
            return new ImportDefinition<XlsxProduct>
            {
                x => x.Name,
                x => x.Id, x => x.Sku, x => x.CategoryPath, x => x.CategoryId, x => x.MainProductId, x => x.PrimaryImage,
                x => x.IsActive, x => x.IsBuyable, x => x.TrackInventory,
                x => x.ManufacturerPartNumber, x => x.Gtin, x => x.MeasureUnit, x => x.WeightUnit, x => x.Weight,
                x => x.Height, x => x.Length, x => x.Width, x => x.TaxType, x => x.ProductType, x => x.ShippingType,
                x => x.Vendor, x => x.DownloadType, x => x.DownloadExpiration, x => x.HasUserAgreement
            };
        }
    }
}