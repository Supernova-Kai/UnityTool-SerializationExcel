//
// Excel工具 工具函数类
//
using UnityEngine;
using UnityEditor;
using System;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml; //OfficeOpenXml需要EPPlus.dll
using System.Xml.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

public class ExcelTool
{
    private static string assetPath = "Assets/Excel/Script/ExcelPathConfig.asset"; //ExcelPathConfig.asset的路径
    public static ExcelPathConfig pathConfig = AssetDatabase.LoadAssetAtPath<ExcelPathConfig>(assetPath); //加载asset文件

    /// <summary>
    /// 将所有Excel文件的"文件信息"添加到List中
    /// </summary>
    public static void GetAllExcelFilesPath(List<FileData> fileList)
    {
        fileList.Clear();
        string[] excelFilesPath = Directory.GetFiles(pathConfig.openPath, "*.xlsx", SearchOption.AllDirectories); //返回所有拓展名为.xlsx的文件的路径
        foreach(string str in excelFilesPath)
        {
            FileInfo fileInfo = new FileInfo(str);
            string excelFileName = Path.GetFileNameWithoutExtension(str); //获取没有拓展名的Excel文件名
            fileList.Add(new FileData(excelFileName, str, false));
        }
    }

    /// <summary>
    /// 根据Excel表格生成数据类
    /// </summary>
    /// <param name="fileList"></param>
    public static void CreateDataScript(List<FileData> fileList)
    {
        foreach(FileData fileData in fileList)
        {
            if(fileData.isSelect) //选中的
            {
                string className = fileData.excelFileName; //以Excel文件名为类名
                string scriptPath = pathConfig.classPath + "/" + className + ".cs";

                using(FileStream fs = new FileStream(scriptPath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite)) //写入脚本
                {
                    using(StreamWriter sw = new StreamWriter(fs)) //写入脚本的writer
                    {
                        sw.WriteLine("//" + className + "类");
                        sw.WriteLine();
                        sw.WriteLine("using System;");
                        sw.WriteLine("using System.Collections.Generic;");
                        sw.WriteLine();
                        sw.WriteLine("[Serializable]");
                        sw.WriteLine("public class " + className);
                        sw.WriteLine("{");

                        FileStream openSteam = File.Open(fileData.excelFilePath, FileMode.Open, FileAccess.Read, FileShare.Read); //Excel表格读取

                        using(ExcelPackage excelPackage = new ExcelPackage(openSteam))
                        {
                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1]; //子表的索引是从1开始的
                            int rowNum = worksheet.Dimension.End.Row; //总行数
                            int columnNum = worksheet.Dimension.End.Column; //总列数
                            for(int i = 1; i <= columnNum; i++) //遍历列（每一列的第二行和第三行对应的是字段的类型和名字）
                            {
                                //Cells的索引是从1开始的，[1, 1]代表最左上角那个单元格 注意：[,]中，左边是行，右边是列
                                string memberType = worksheet.Cells[2, i].Value.ToString(); //在我们的约定格式中，每一张表的第二行是字段的类型（这里获取了第二行中每一列的值）
                                string memberName = worksheet.Cells[3, i].Value.ToString(); //第三行是字段类型
                                string memberDes = worksheet.Cells[1, i].Value.ToString(); //第一行是字段的描述（作为注释）
                                sw.WriteLine($"\tpublic {memberType} {memberName}; //{memberDes}"); //写入属性及其注释
                            }
                            sw.WriteLine("}");
                        }

                        sw.WriteLine();
                        sw.WriteLine("[Serializable]");
                        sw.WriteLine($"public class {className}List");
                        sw.WriteLine("{");
                        sw.WriteLine($"\tpublic List<{className}> dataList;");
                        sw.WriteLine("}");
                    }
                }
                AssetDatabase.Refresh(); //刷新资源
            }
        }
    }

    /// <summary>
    /// 读取Excel表格，将数据添加到生成的数据类中
    /// </summary>
    public static void AddDataToList(List<FileData> fileList)
    {
        foreach(FileData fileData in fileList)
        {
            if(fileData.isSelect)
            {
                string className = fileData.excelFileName; //数据类名

                Type listClassType = Type.GetType(className + "List"); //List类的类型
                FieldInfo dataList = listClassType.GetField("dataList"); //反射出dataList属性
                System.Object listClassInstance = Activator.CreateInstance(listClassType); //创建List类的实例
                
                Type dataClass = Type.GetType(className); //Excel数据对应的类的类型
                FieldInfo[] fls = dataClass.GetFields(); //反射出其全部属性

                System.Object dataListInstance = GetDataListInstance(dataClass); //dataList的实例

                FileStream fs = File.Open(fileData.excelFilePath, FileMode.Open, FileAccess.Read, FileShare.Read); //读取Excel表格
                using(ExcelPackage excelPackage = new ExcelPackage(fs))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[1]; //读取第一张子表
                    int rowNum = worksheet.Dimension.End.Row; //总行数
                    int columnNum = worksheet.Dimension.End.Column; //总列数

                    for(int i = 4; i <= rowNum; i++) //第四行开始才是真正的数据（索引是从1开始的）
                    {
                        System.Object item = Activator.CreateInstance(dataClass); //每一行创建一个实例
                        for(int j = 1; j <= columnNum; j++) //遍历列
                        {
                            string valueType = worksheet.Cells[2, j].Value.ToString(); //字段的类型（表中第二行）
                            string valueName = worksheet.Cells[3, j].Value.ToString(); //字段名字（表中第三行）
                            string value = worksheet.Cells[i, j].Value.ToString(); //数据
                            FieldInfo valueFl = GetFieldInfoByName(fls, valueName);
                            if(valueType == "int") //int类型
                            {
                                valueFl.SetValue(item, Convert.ToInt32(value));
                            }
                            else if(valueType == "float") //float类型
                            {
                                valueFl.SetValue(item, Convert.ToSingle(value));
                            }
                            else if(valueType == "double") //double类型
                            {
                                valueFl.SetValue(item, Convert.ToDouble(value));
                            }
                            else if(valueType == "bool") //bool类型
                            {
                                valueFl.SetValue(item, Convert.ToBoolean(value));
                            }
                            else if(valueType == "string") //string类型
                            {
                                valueFl.SetValue(item, value);
                            }
                        }
                        AddItem(ref dataListInstance, item);
                    }
                    dataList.SetValue(listClassInstance, dataListInstance);
                }
                fs.Close();
                SerializationToJson(listClassInstance, className);
                SerializationToXml(listClassInstance, className);
                SerializationToBinary(listClassInstance, className);
                AssetDatabase.Refresh();
            }
        }
    }

    /// <summary>
    /// 通过名字去获取到对面的属性
    /// </summary>
    /// <param name="fls">数据类的索引属性（数组）</param>
    /// <param name="name">名字</param>
    /// <returns></returns>
    private static FieldInfo GetFieldInfoByName(FieldInfo[] fls, string name)
    {
        FieldInfo fl = null;
        for(int i = 0; i < fls.Length; i++)
        {
            if(fls[i].Name == name)
            {
                fl = fls[i];
            }
        }
        return fl;
    }

    /// <summary>
    /// 获取dataList的实例
    /// </summary>
    /// <param name="t">集合元素的类型</param>
    /// <returns></returns>
    private static object GetDataListInstance(Type t)
    {
        Type listType = typeof(List<>); //List类型
        Type listBaseType = t; //集合元素的类型
        Type listClassType = listType.MakeGenericType(new System.Type[] { listBaseType });
        return Activator.CreateInstance(listClassType, new object[] { });
    }

    /// <summary>
    /// 将每一条数据添加到dataList中
    /// </summary>
    /// <param name="dataList">存储数据的集合（即dataList）</param>
    /// <param name="item">每一条数据</param>
    private static void AddItem(ref object dataList, object item)
    {
        dataList.GetType().InvokeMember("Add", BindingFlags.Default | BindingFlags.InvokeMethod, null, dataList, new object[] { item });
    }

    /// <summary>
    /// 序列化为Json文件
    /// </summary>
    /// <param name="listClass">序列化对象（即数据的list类）</param>
    /// <param name="className">数据类名（作为文件名）</param>
    private static void SerializationToJson(System.Object listClass, string className)
    {
        string json = JsonUtility.ToJson(listClass, true);
        File.WriteAllText(pathConfig.jsonPath + "/" + className + ".json", json);
    }

    /// <summary>
    /// 序列化为Xml文件
    /// </summary>
    /// <param name="listClass">序列化对象（即数据的list类）</param>
    /// <param name="className">数据类名（作为文件名）</param>
    private static void SerializationToXml(System.Object listClass, string className)
    {
        string filePath = pathConfig.xmlPath + "/" + className + ".xml";
        using(FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            using(StreamWriter sw = new StreamWriter(fs))
            {
                XmlSerializer xs = new XmlSerializer(listClass.GetType());
                xs.Serialize(sw, listClass);
            }
        }
    }

    /// <summary>
    /// 序列化为二进制文件
    /// </summary>
    /// <param name="listClass">序列化对象（即数据的list类）</param>
    /// <param name="className">数据类名（作为文件名）</param>
    private static void SerializationToBinary(System.Object listClass, string className)
    {
        string filePath = pathConfig.binaryPath + "/" + className;
        using(FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            BinaryFormatter bf = new BinaryFormatter();
            bf.Serialize(fs, listClass);
        }
    }
}