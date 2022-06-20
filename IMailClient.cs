namespace PrintMailExcel
{
    public interface IMailClient
    {
        string EmailAddress { get; }
        string Password { get; }
        bool TryGetExcelFilesInfo(string basePath, out List<ExcelFile> excelFilesInfo);
        List<ExcelPrint> GetRows(List<ExcelFile> excelFiles);
    }
}