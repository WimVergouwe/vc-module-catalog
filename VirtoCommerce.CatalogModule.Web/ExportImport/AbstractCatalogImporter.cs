using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using VirtoCommerce.CatalogModule.Data.Repositories;
using VirtoCommerce.Domain.Catalog.Model;
using VirtoCommerce.Domain.Catalog.Services;
using VirtoCommerce.Domain.Commerce.Services;
using VirtoCommerce.Domain.Inventory.Services;
using VirtoCommerce.Domain.Pricing.Services;
using VirtoCommerce.Platform.Core.Common;
using VirtoCommerce.Platform.Core.ExportImport;

namespace VirtoCommerce.CatalogModule.Web.ExportImport
{
    public abstract class AbstractCatalogImporter
    {
        private readonly char[] _categoryDelimiters = { '/', '|', '\\', '>' };
        private readonly ICategoryService _categoryService;
        private readonly IItemService _productService;
        private readonly ISkuGenerator _skuGenerator;
        private readonly IPricingService _pricingService;
        private readonly IInventoryService _inventoryService;
        private readonly ICommerceService _commerceService;
        private readonly IPropertyService _propertyService;
        private readonly ICatalogSearchService _searchService;
        private readonly Func<ICatalogRepository> _catalogRepositoryFactory;
        private readonly object _lockObject = new object();


        protected AbstractCatalogImporter(ICategoryService categoryService, IItemService productService,
            ISkuGenerator skuGenerator,
            IPricingService pricingService, IInventoryService inventoryService, ICommerceService commerceService,
            IPropertyService propertyService, ICatalogSearchService searchService, Func<ICatalogRepository> catalogRepositoryFactory)
        {
            _categoryService = categoryService;
            _productService = productService;
            _skuGenerator = skuGenerator;
            _pricingService = pricingService;
            _inventoryService = inventoryService;
            _commerceService = commerceService;
            _propertyService = propertyService;
            _searchService = searchService;
            _catalogRepositoryFactory = catalogRepositoryFactory;
        }

        public abstract void DoImport(Stream inputStream, ImportInfo importInfo, Action<ExportImportProgressInfo> progressCallback);

        protected void SaveCategoryTree(Catalog catalog, IEnumerable<CatalogProduct> catalogProducts, ExportImportProgressInfo progressInfo, Action<ExportImportProgressInfo> progressCallback)
        {
            progressInfo.ProcessedCount = 0;
            var cachedCategoryMap = new Dictionary<string, Category>();

            foreach (var catalogProduct in catalogProducts.Where(x => x.Category != null && !string.IsNullOrEmpty(x.Category.Path)))
            {
                var outline = "";
                var productCategoryNames = catalogProduct.Category.Path.Split(_categoryDelimiters);
                string parentCategoryId = null;
                foreach (var categoryName in productCategoryNames)
                {
                    outline += "\\" + categoryName;
                    Category category;
                    if (!cachedCategoryMap.TryGetValue(outline, out category))
                    {
                        var searchCriteria = new SearchCriteria
                        {
                            CatalogId = catalog.Id,
                            CategoryId = parentCategoryId,
                            ResponseGroup = SearchResponseGroup.WithCategories
                        };
                        category = _searchService.Search(searchCriteria).Categories.FirstOrDefault(x => x.Name == categoryName);
                    }

                    if (category == null)
                    {
                        category = _categoryService.Create(new Category() { Name = categoryName, Code = categoryName.GenerateSlug(), CatalogId = catalog.Id, ParentId = parentCategoryId });
                        //Raise notification each notifyCategorySizeLimit category
                        progressInfo.Description = string.Format("Creating categories: {0} created", ++progressInfo.ProcessedCount);
                        progressCallback(progressInfo);
                    }
                    catalogProduct.CategoryId = category.Id;
                    catalogProduct.Category = category;
                    parentCategoryId = category.Id;
                    cachedCategoryMap[outline] = category;
                }
            }
        }

        protected void SaveProducts(Catalog catalog, List<CatalogProduct> catalogProducts, ExportImportProgressInfo progressInfo, Action<ExportImportProgressInfo> progressCallback)
        {
            progressInfo.ProcessedCount = 0;
            progressInfo.TotalCount = catalogProducts.Count;

            var defaultFulfilmentCenter = _commerceService.GetAllFulfillmentCenters().FirstOrDefault();
            DetectParents(catalogProducts);

            //Detect already exist product by Code
            using (var repository = _catalogRepositoryFactory())
            {
                var codes = catalogProducts.Where(x => x.IsTransient()).Select(x => x.Code).Where(x => x != null).Distinct().ToArray();
                var existProducts = repository.Items.Where(x => x.CatalogId == catalog.Id && codes.Contains(x.Code)).Select(x => new { Id = x.Id, Code = x.Code }).ToArray();
                foreach (var existProduct in existProducts)
                {
                    var product = catalogProducts.FirstOrDefault(x => x.Code == existProduct.Code);
                    if (product != null)
                    {
                        product.Id = product.Id;
                    }
                }
            }

            var categoriesIds = catalogProducts.Where(x => x.CategoryId != null).Select(x => x.CategoryId).Distinct().ToArray();
            var categpories = _categoryService.GetByIds(categoriesIds, CategoryResponseGroup.WithProperties);

            var defaultLanguge = catalog.DefaultLanguage != null ? catalog.DefaultLanguage.LanguageCode : "EN-US";
            var changedProperties = new List<Property>();
            foreach (var catalogProduct in catalogProducts)
            {
                catalogProduct.CatalogId = catalog.Id;
                if (string.IsNullOrEmpty(catalogProduct.Code))
                {
                    catalogProduct.Code = _skuGenerator.GenerateSku(catalogProduct);
                }
                //Set a parent relations
                if (catalogProduct.MainProductId == null && catalogProduct.MainProduct != null)
                {
                    catalogProduct.MainProductId = catalogProduct.MainProduct.Id;
                }
                var csvProduct = catalogProduct as CsvProduct;
                if (csvProduct != null)
                {
                    csvProduct.EditorialReview.LanguageCode = defaultLanguge;
                    csvProduct.SeoInfo.LanguageCode = defaultLanguge;
                    csvProduct.SeoInfo.SemanticUrl = string.IsNullOrEmpty(csvProduct.SeoInfo.SemanticUrl) ? catalogProduct.Code : csvProduct.SeoInfo.SemanticUrl;
                }

                var properties = catalog.Properties;
                if (catalogProduct.CategoryId != null)
                {
                    var category = categpories.FirstOrDefault(x => x.Id == catalogProduct.CategoryId);
                    if (category != null)
                    {
                        properties = category.Properties;
                    }
                }

                //Try to fill properties meta information for values
                foreach (var propertyValue in catalogProduct.PropertyValues)
                {
                    if (propertyValue.Value != null)
                    {
                        var property = properties.FirstOrDefault(x => string.Equals(x.Name, propertyValue.PropertyName));
                        if (property != null)
                        {
                            propertyValue.ValueType = property.ValueType;
                            if (property.Dictionary)
                            {
                                var dicValue = property.DictionaryValues.FirstOrDefault(x => Equals(x.Value, propertyValue.Value));
                                if (dicValue == null)
                                {
                                    dicValue = new PropertyDictionaryValue
                                    {
                                        Alias = propertyValue.Value.ToString(),
                                        Value = propertyValue.Value.ToString(),
                                        Id = Guid.NewGuid().ToString()
                                    };
                                    property.DictionaryValues.Add(dicValue);
                                    if (!changedProperties.Contains(property))
                                    {
                                        changedProperties.Add(property);
                                    }
                                }
                                propertyValue.ValueId = dicValue.Id;
                            }
                        }
                    }
                }
            }

            progressInfo.Description = string.Format("Saving property dictionary values...");
            progressCallback(progressInfo);
            _propertyService.Update(changedProperties.ToArray());

            var options = new ParallelOptions
            {
                MaxDegreeOfParallelism = 10
            };

            foreach (var group in new[] { catalogProducts.Where(x => x.MainProduct == null), catalogProducts.Where(x => x.MainProduct != null) })
            {
                var partSize = 25;
                var partsCount = Math.Max(1, group.Count() / partSize);
                var parts = group.Select((x, i) => new { Index = i, Value = x })
                    .GroupBy(x => x.Index % partsCount)
                    .Select(x => x.Select(v => v.Value));

                Parallel.ForEach(parts, options, products =>
                {
                    try
                    {
                        //Save main products first and then variations
                        var toUpdateProducts = products.Where(x => !x.IsTransient());
                        //Need to additional  check that  product with id exist  
                        using (var repository = _catalogRepositoryFactory())
                        {
                            var updateProductIds = toUpdateProducts.Select(x => x.Id).ToArray();
                            var existProductIds = repository.Items.Where(x => updateProductIds.Contains(x.Id)).Select(x => x.Id).ToArray();
                            toUpdateProducts = toUpdateProducts.Where(x => existProductIds.Contains(x.Id)).ToList();
                        }
                        var toCreateProducts = products.Except(toUpdateProducts);
                        if (!toCreateProducts.IsNullOrEmpty())
                        {
                            _productService.Create(toCreateProducts.ToArray());
                        }
                        if (!toUpdateProducts.IsNullOrEmpty())
                        {
                            _productService.Update(toUpdateProducts.ToArray());
                        }


                        var csvProducts = products.Select(x => x as CsvProduct).Where(x => x != null).ToList();
                        //Set productId for dependent objects
                        foreach (var product in csvProducts)
                        {
                            product.Inventory.ProductId = product.Id;
                            product.Inventory.FulfillmentCenterId = product.Inventory.FulfillmentCenterId ?? defaultFulfilmentCenter.Id;
                            product.Price.ProductId = product.Id;
                        }

                        var inventories = csvProducts.Where(x => x.Inventory != null).Select(x => x.Inventory).ToArray();
                        _inventoryService.UpsertInventories(inventories);

                        var prices = csvProducts.Where(x => x.Price != null && x.Price.EffectiveValue > 0).Select(x => x.Price).ToArray();
                        _pricingService.SavePrices(prices);


                    }
                    catch (Exception ex)
                    {
                        lock (_lockObject)
                        {
                            progressInfo.Errors.Add(ex.ToString());
                            progressCallback(progressInfo);
                        }
                    }
                    finally
                    {
                        lock (_lockObject)
                        {
                            //Raise notification
                            progressInfo.ProcessedCount += products.Count();
                            progressInfo.Description = string.Format("Saving products: {0} of {1} created", progressInfo.ProcessedCount, progressInfo.TotalCount);
                            progressCallback(progressInfo);
                        }
                    }
                });
            }
        }

        protected void DetectParents(List<CatalogProduct> catalogProducts)
        {
            foreach (var catalogProduct in catalogProducts)
            {
                //Try to set parent relations
                //By id or code reference
                var parentProduct = catalogProducts.FirstOrDefault(x => catalogProduct.MainProductId != null && (x.Id == catalogProduct.MainProductId || x.Code == catalogProduct.MainProductId));
                catalogProduct.MainProduct = parentProduct;
                catalogProduct.MainProductId = parentProduct != null ? parentProduct.Id : null;
            }
        }
    }
}