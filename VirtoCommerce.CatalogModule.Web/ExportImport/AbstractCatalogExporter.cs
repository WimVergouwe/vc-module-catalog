using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using VirtoCommerce.Domain.Catalog.Model;
using VirtoCommerce.Domain.Catalog.Services;
using VirtoCommerce.Platform.Core.Assets;
using VirtoCommerce.Platform.Core.ExportImport;

namespace VirtoCommerce.CatalogModule.Web.ExportImport
{
    public abstract class AbstractCatalogExporter
    {
        private readonly ICatalogSearchService _searchService;
        private readonly IItemService _productService;
        protected readonly IBlobUrlResolver BlobUrlResolver;

        protected AbstractCatalogExporter(ICatalogSearchService catalogSearchService, IItemService productService, IBlobUrlResolver blobUrlResolver)
        {
            _searchService = catalogSearchService;
            _productService = productService;
            BlobUrlResolver = blobUrlResolver;
        }

        public abstract void DoExport(Stream outStream, ExportInfo exportInfo, Action<ExportImportProgressInfo> progressCallback);

        protected List<CatalogProduct> LoadProducts(string catalogId, string[] exportedCategories, string[] exportedProducts)
        {
            var retVal = new List<CatalogProduct>();

            var productIds = new List<string>();
            if (exportedProducts != null)
            {
                productIds = exportedProducts.ToList();
            }
            if (exportedCategories != null && exportedCategories.Any())
            {
                foreach (var categoryId in exportedCategories)
                {
                    var result = _searchService.Search(new SearchCriteria { CatalogId = catalogId, CategoryId = categoryId, Skip = 0, Take = int.MaxValue, ResponseGroup = SearchResponseGroup.WithProducts | SearchResponseGroup.WithCategories });
                    productIds.AddRange(result.Products.Select(x => x.Id));
                    if (result.Categories != null && result.Categories.Any())
                    {
                        retVal.AddRange(LoadProducts(catalogId, result.Categories.Select(x => x.Id).ToArray(), null));
                    }
                }
            }

            if ((exportedCategories == null || !exportedCategories.Any()) && (exportedProducts == null || !exportedProducts.Any()))
            {
                var result = _searchService.Search(new SearchCriteria { CatalogId = catalogId, SearchInChildren = true, Skip = 0, Take = int.MaxValue, ResponseGroup = SearchResponseGroup.WithProducts });
                productIds = result.Products.Select(x => x.Id).ToList();
            }

            var products = _productService.GetByIds(productIds.Distinct().ToArray(), ItemResponseGroup.ItemLarge);
            foreach (var product in products)
            {
                retVal.Add(product);
                if (product.Variations != null)
                {
                    retVal.AddRange(product.Variations);
                }
            }

            return retVal;
        }
    }
}