//
// ����ExcelPathConfig.asset ���ڹ������·��
//
using System;
using System.IO;
using UnityEngine;

[CreateAssetMenu(fileName = "ExcelPathConfig", menuName = "Create Excel Path Config", order = 0)]
[Serializable]
public class ExcelPathConfig : ScriptableObject
{
    [Header("Excel�ļ���·��")]
    [ContextMenuItem("Default", "DefaultOpenPath")]
    public string openPath = string.Empty; //Excel�ļ���·��
    private void DefaultOpenPath() //����openPathĬ��ֵ
    {
        openPath = Application.dataPath + "/Excel/Data/Input";
        if(!Directory.Exists(openPath))
        {
            Directory.CreateDirectory(openPath);
        }
    }

    [Header("���ݽű����·��")]
    [Tooltip("�����߻����Excel����Զ��������Ӧ��������ű�����·��Ϊ���ݽű������·��")]
    [ContextMenuItem("Default", "DefaultClassPath")]
    public string classPath = string.Empty; //���ݽű����·��
    private void DefaultClassPath() //����classPathĬ��ֵ
    {
        classPath = Application.dataPath + "/Excel/Data/Output/Class";
        if(!Directory.Exists(classPath))
        {
            Directory.CreateDirectory(classPath);
        }
    }

    [Header("Json�ļ����·��")]
    [ContextMenuItem("Default", "DefaultJsonPath")]
    public string jsonPath = string.Empty; //Json�ļ����·��
    private void DefaultJsonPath() //����jsonPathĬ��ֵ
    {
        jsonPath = Application.dataPath + "/Excel/Data/Output/Json";
        if(!Directory.Exists(jsonPath))
        {
            Directory.CreateDirectory(jsonPath);
        }
    }

    [Header("Xml�ļ����·��")]
    [ContextMenuItem("Default", "DefaultXmlPath")]
    public string xmlPath = string.Empty; //Xml�ļ����·��
    private void DefaultXmlPath() //����xmlPathĬ��ֵ
    {
        xmlPath = Application.dataPath + "/Excel/Data/Output/Xml";
        if(!Directory.Exists(xmlPath))
        {
            Directory.CreateDirectory(xmlPath);
        }
    }

    [Header("�������ļ����·��")]
    [ContextMenuItem("Default", "DefaultBinaryPath")]
    public string binaryPath = string.Empty; //�������ļ����·��
    private void DefaultBinaryPath() //����binaryPathĬ��ֵ
    {
        binaryPath = Application.dataPath + "/Excel/Data/Output/Binary";
        if(!Directory.Exists(binaryPath))
        {
            Directory.CreateDirectory(binaryPath);
        }
    }
}