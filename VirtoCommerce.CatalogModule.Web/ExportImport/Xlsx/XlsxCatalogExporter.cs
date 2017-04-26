using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using VirtoCommerce.CatalogModule.Web.Utilities;
using VirtoCommerce.Domain.Catalog.Model;
using VirtoCommerce.Domain.Catalog.Services;
using VirtoCommerce.Platform.Core.ExportImport;

namespace VirtoCommerce.CatalogModule.Web.ExportImport.Xlsx
{
    public class XlsxCatalogExporter : AbstractCatalogExporter
    {
        public XlsxCatalogExporter(ICatalogSearchService catalogSearchService, IItemService productService) : base(catalogSearchService, productService)
        {
        }

        public override void DoExport(Stream outStream, ExportInfo exportInfo, Action<ExportImportProgressInfo> progressCallback)
        {
            var progressInfo = new ExportImportProgressInfo
            {
                Description = "loading products..."
            };

            //Notification
            progressCallback(progressInfo);

            //Load all products to export
            var products = LoadProducts(exportInfo.CatalogId, exportInfo.CategoryIds, exportInfo.ProductIds);

            DemoXlsxUtilities.CreateFileWithData(outStream, "CatalogExport", GetExportDefinition(products), products);
        }

        private ExportDefinition<CatalogProduct> GetExportDefinition(IEnumerable<CatalogProduct> products)
        {
            var definition = new ExportDefinition<CatalogProduct>
            {
                x => x.Id, x => x.CategoryId, x => x.MainProductId, x => x.IsActive, x => x.IsBuyable, x => x.TrackInventory,
                x => x.ManufacturerPartNumber, x => x.Gtin, x => x.MeasureUnit, x => x.WeightUnit, x => x.Weight,
                x => x.Height, x => x.Length, x => x.Width, x => x.TaxType, x => x.ProductType, x => x.ShippingType,
                x => x.Vendor, x => x.DownloadType, x => x.DownloadExpiration, x => x.HasUserAgreement
            };

            foreach (var propertyValue in products.SelectMany(product => product.PropertyValues))
            {
                definition.Add(new ColumnExportDefinition<CatalogProduct>(propertyValue.PropertyName, p =>
                {
                    return p.PropertyValues.Where(x => x.PropertyName.Equals(propertyValue.PropertyName, StringComparison.OrdinalIgnoreCase));
                }));
            }

            return definition;
        }
    }
}