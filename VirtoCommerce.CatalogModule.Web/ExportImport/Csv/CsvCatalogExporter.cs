using System;
using System.IO;
using System.Linq;
using System.Text;
using CsvHelper;
using VirtoCommerce.Domain.Catalog.Services;
using VirtoCommerce.Domain.Inventory.Services;
using VirtoCommerce.Domain.Pricing.Model;
using VirtoCommerce.Domain.Pricing.Services;
using VirtoCommerce.Platform.Core.Assets;
using VirtoCommerce.Platform.Core.ExportImport;

namespace VirtoCommerce.CatalogModule.Web.ExportImport.Csv
{
    public sealed class CsvCatalogExporter : AbstractCatalogExporter
    {
        private readonly IPricingService _pricingService;
        private readonly IInventoryService _inventoryService;
        private readonly IBlobUrlResolver _blobUrlResolver;

        public CsvCatalogExporter(ICatalogSearchService catalogSearchService, IItemService productService,
            IPricingService pricingService, IInventoryService inventoryService, IBlobUrlResolver blobUrlResolver)
            : base(catalogSearchService, productService)
        {
            _pricingService = pricingService;
            _inventoryService = inventoryService;
            _blobUrlResolver = blobUrlResolver;
        }

        public override void DoExport(Stream outStream, ExportInfo exportInfo, Action<ExportImportProgressInfo> progressCallback)
        {
            var prodgressInfo = new ExportImportProgressInfo
            {
                Description = "loading products..."
            };

            var streamWriter = new StreamWriter(outStream, Encoding.UTF8, 1024, true) { AutoFlush = true };
            using (var csvWriter = new CsvWriter(streamWriter))
            {
                //Notification
                progressCallback(prodgressInfo);

                //Load all products to export
                var products = LoadProducts(exportInfo.CatalogId, exportInfo.CategoryIds, exportInfo.ProductIds);
                var allProductIds = products.Select(x => x.Id).ToArray();

                //Load prices for products
                prodgressInfo.Description = "loading prices...";
                progressCallback(prodgressInfo);

                var priceEvalContext = new PriceEvaluationContext
                {
                    ProductIds = allProductIds,
                    PricelistIds = exportInfo.PriceListId == null ? null : new[] { exportInfo.PriceListId },
                    Currency = exportInfo.Currency
                };
                var allProductPrices = _pricingService.EvaluateProductPrices(priceEvalContext).ToList();


                //Load inventories
                prodgressInfo.Description = "loading inventory information...";
                progressCallback(prodgressInfo);

                var allProductInventories = _inventoryService.GetProductsInventoryInfos(allProductIds).Where(x => exportInfo.FulfilmentCenterId == null || x.FulfillmentCenterId == exportInfo.FulfilmentCenterId).ToList();


                //Export configuration
                var csvProductMappingConfiguration = CsvProductMappingConfiguration.GetDefaultConfiguration();
                csvProductMappingConfiguration.PropertyCsvColumns = products.SelectMany(x => x.PropertyValues).Select(x => x.PropertyName).Distinct().ToArray();

                csvWriter.Configuration.Delimiter = csvProductMappingConfiguration.Delimiter;
                csvWriter.Configuration.RegisterClassMap(new CsvProductMap(csvProductMappingConfiguration));

                //Write header
                csvWriter.WriteHeader<CsvProduct>();

                prodgressInfo.TotalCount = products.Count;
                var notifyProductSizeLimit = 50;
                var counter = 0;
                foreach (var product in products)
                {
                    try
                    {
                        var csvProduct = new CsvProduct(product, _blobUrlResolver, allProductPrices.FirstOrDefault(x => x.ProductId == product.Id), allProductInventories.FirstOrDefault(x => x.ProductId == product.Id));
                        csvWriter.WriteRecord(csvProduct);
                    }
                    catch (Exception ex)
                    {
                        prodgressInfo.Errors.Add(ex.ToString());
                        progressCallback(prodgressInfo);
                    }

                    //Raise notification each notifyProductSizeLimit products
                    counter++;
                    prodgressInfo.ProcessedCount = counter;
                    prodgressInfo.Description = $"{prodgressInfo.ProcessedCount} of {prodgressInfo.TotalCount} products processed";
                    if (counter % notifyProductSizeLimit == 0 || counter == prodgressInfo.TotalCount)
                    {
                        progressCallback(prodgressInfo);
                    }
                }
            }
        }
    }
}
