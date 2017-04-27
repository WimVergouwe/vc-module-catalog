using System.Collections.Generic;
using System.Linq;
using Omu.ValueInjecter;
using VirtoCommerce.Domain.Catalog.Model;
using VirtoCommerce.Domain.Commerce.Model;
using VirtoCommerce.Domain.Inventory.Model;
using VirtoCommerce.Domain.Pricing.Model;
using VirtoCommerce.Platform.Core.Assets;

namespace VirtoCommerce.CatalogModule.Web.ExportImport
{
    public class XlsxProduct : CatalogProduct
    {
        private readonly IBlobUrlResolver _blobUrlResolver;
        public XlsxProduct()
        {
            Properties = new List<Property>();
            PropertyValues = new List<PropertyValue>();
            Images = new List<Image>();
            Assets = new List<Asset>();
            Links = new List<CategoryLink>();
            Variations = new List<CatalogProduct>();
            SeoInfos = new List<SeoInfo>();
            Reviews = new List<EditorialReview>();
            Associations = new List<ProductAssociation>();
            Prices = new List<Price>();
            Inventories = new List<InventoryInfo>();
            Outlines = new List<Outline>();
        }

        public XlsxProduct(CatalogProduct product, IBlobUrlResolver blobUrlResolver)
            : this()
        {
            _blobUrlResolver = blobUrlResolver;

            this.InjectFrom(product);
            PropertyValues = product.PropertyValues;
            Images = product.Images;
            Assets = product.Assets;
            Links = product.Links;
            Variations = product.Variations;
            SeoInfos = product.SeoInfos;
            Reviews = product.Reviews;
            Associations = product.Associations;
        }

        public string Sku
        {
            get { return Code; }
            set { Code = value; }
        }

        public string CategoryPath
        {
            get
            {
                return Category != null ? string.Join("/", Category.Parents.Select(x => x.Name).Concat(new[] { Category.Name })) : null;
            }
            set
            {
                Category = new Category { Path = value };
            }
        }

        public string PrimaryImage
        {
            get
            {
                var retVal = string.Empty;
                if (Images != null)
                {
                    var primaryImage = Images.OrderBy(x => x.SortOrder).FirstOrDefault();
                    if (primaryImage != null)
                    {
                        retVal = _blobUrlResolver != null ? _blobUrlResolver.GetAbsoluteUrl(primaryImage.Url) : primaryImage.Url;
                    }
                }
                return retVal;
            }

            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    Images.Add(new Image
                    {
                        Url = value,
                        SortOrder = 0
                    });
                }
            }
        }
    }
}
