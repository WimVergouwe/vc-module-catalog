using System;
using System.IO;
using VirtoCommerce.CatalogModule.Web.Utilities;
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

            DemoXlsxUtilities.CreateFileWithData(outStream, "CatalogExport", products);
        }
    }
}