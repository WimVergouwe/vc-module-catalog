namespace VirtoCommerce.CatalogModule.Web.ExportImport
{
    public class ExportInfo
    {

        public string CatalogId { get; set; }
        public string[] ProductIds { get; set; }
        public string[] CategoryIds { get; set; }
        public string PriceListId { get; set; }
        public string FulfilmentCenterId { get; set; }
        public string Currency { get; set; }
    }

    public class CsvExportInfo : ExportInfo
    {
        public CsvProductMappingConfiguration Configuration { get; set; }
    }
}
