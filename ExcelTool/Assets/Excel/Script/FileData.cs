/// <summary>
/// FileData类，用于管理文件
/// </summary>
public class FileData
{
    /// <summary>
    /// Excel文件 文件名
    /// </summary>
    public string excelFileName;

    /// <summary>
    /// Excel文件 路径
    /// </summary>
    public string excelFilePath;

    /// <summary>
    /// 在工具界面是否选中
    /// </summary>
    public bool isSelect;

    /// <summary>
    /// FileData构造方法
    /// </summary>
    /// <param name="_excelFileName">Excel文件 文件名</param>
    /// <param name="_excelFilePath">Excel文件 路径</param>
    /// <param name="_isSelect">在工具界面是否选中</param>
    public FileData(string _excelFileName, string _excelFilePath, bool _isSelect)
    {
        this.excelFileName = _excelFileName;
        this.excelFilePath = _excelFilePath;
        this.isSelect = _isSelect;
    }
}