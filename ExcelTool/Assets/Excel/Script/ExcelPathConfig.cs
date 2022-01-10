//
// 创建ExcelPathConfig.asset 便于管理相关路径
//
using System;
using System.IO;
using UnityEngine;

[CreateAssetMenu(fileName = "ExcelPathConfig", menuName = "Create Excel Path Config", order = 0)]
[Serializable]
public class ExcelPathConfig : ScriptableObject
{
    [Header("Excel文件打开路径")]
    [ContextMenuItem("Default", "DefaultOpenPath")]
    public string openPath = string.Empty; //Excel文件打开路径
    private void DefaultOpenPath() //设置openPath默认值
    {
        openPath = Application.dataPath + "/Excel/Data/Input";
        if(!Directory.Exists(openPath))
        {
            Directory.CreateDirectory(openPath);
        }
    }

    [Header("数据脚本输出路径")]
    [Tooltip("本工具会根据Excel表格，自动生成相对应的数据类脚本，此路径为数据脚本的输出路径")]
    [ContextMenuItem("Default", "DefaultClassPath")]
    public string classPath = string.Empty; //数据脚本输出路径
    private void DefaultClassPath() //设置classPath默认值
    {
        classPath = Application.dataPath + "/Excel/Data/Output/Class";
        if(!Directory.Exists(classPath))
        {
            Directory.CreateDirectory(classPath);
        }
    }

    [Header("Json文件输出路径")]
    [ContextMenuItem("Default", "DefaultJsonPath")]
    public string jsonPath = string.Empty; //Json文件输出路径
    private void DefaultJsonPath() //设置jsonPath默认值
    {
        jsonPath = Application.dataPath + "/Excel/Data/Output/Json";
        if(!Directory.Exists(jsonPath))
        {
            Directory.CreateDirectory(jsonPath);
        }
    }

    [Header("Xml文件输出路径")]
    [ContextMenuItem("Default", "DefaultXmlPath")]
    public string xmlPath = string.Empty; //Xml文件输出路径
    private void DefaultXmlPath() //设置xmlPath默认值
    {
        xmlPath = Application.dataPath + "/Excel/Data/Output/Xml";
        if(!Directory.Exists(xmlPath))
        {
            Directory.CreateDirectory(xmlPath);
        }
    }

    [Header("二进制文件输出路径")]
    [ContextMenuItem("Default", "DefaultBinaryPath")]
    public string binaryPath = string.Empty; //二进制文件输出路径
    private void DefaultBinaryPath() //设置binaryPath默认值
    {
        binaryPath = Application.dataPath + "/Excel/Data/Output/Binary";
        if(!Directory.Exists(binaryPath))
        {
            Directory.CreateDirectory(binaryPath);
        }
    }
}