namespace VirtoCommerce.CatalogModule.Web.ExportImport
{
    public class ImportInfo
    {
        public string CatalogId { get; set; }
        public string FileUrl { get; set; }
    }
    public class CsvImportInfo : ImportInfo {
        public CsvProductMappingConfiguration Configuration { get; set; }
    }
}
