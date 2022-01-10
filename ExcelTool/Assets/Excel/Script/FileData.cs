/// <summary>
/// FileData�࣬���ڹ����ļ�
/// </summary>
public class FileData
{
    /// <summary>
    /// Excel�ļ� �ļ���
    /// </summary>
    public string excelFileName;

    /// <summary>
    /// Excel�ļ� ·��
    /// </summary>
    public string excelFilePath;

    /// <summary>
    /// �ڹ��߽����Ƿ�ѡ��
    /// </summary>
    public bool isSelect;

    /// <summary>
    /// FileData���췽��
    /// </summary>
    /// <param name="_excelFileName">Excel�ļ� �ļ���</param>
    /// <param name="_excelFilePath">Excel�ļ� ·��</param>
    /// <param name="_isSelect">�ڹ��߽����Ƿ�ѡ��</param>
    public FileData(string _excelFileName, string _excelFilePath, bool _isSelect)
    {
        this.excelFileName = _excelFileName;
        this.excelFilePath = _excelFilePath;
        this.isSelect = _isSelect;
    }
}