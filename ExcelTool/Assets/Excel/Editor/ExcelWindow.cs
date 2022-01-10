//
// Excel工具UI部分
//
using UnityEngine;
using UnityEditor;
using System.Collections.Generic;
using System.IO;

public class ExcelWindow : EditorWindow
{
    private static ExcelWindow excelWindow;
    public List<FileData> excelFilesList = new List<FileData>();

    [MenuItem("Tools/Excel工具", false, -10)]
    static void ShowWindow()
    {
        excelWindow = EditorWindow.GetWindow<ExcelWindow>();
        excelWindow.titleContent = new GUIContent("Excel工具");
        excelWindow.position = new Rect(200, 300, 800, 400);
    }

    private void OnGUI()
    {
        GUILayout.Label("Excel工具", "WarningOverlay");
        if(GUILayout.Button("读取路径下的Excel文件"))
        {
            ExcelTool.GetAllExcelFilesPath(excelFilesList);
        }
        GUILayout.BeginVertical("box");
        foreach(FileData fileData in excelFilesList)
        {
            GUILayout.BeginHorizontal();
            fileData.isSelect = EditorGUILayout.Toggle(fileData.isSelect, GUILayout.Width(20));
            GUILayout.Label("文件名：" + fileData.excelFileName);
            GUILayout.Space(30);
            GUILayout.Label("路径" + fileData.excelFilePath);
            GUILayout.EndHorizontal();
            GUILayout.Space(5);
        }
        GUILayout.EndVertical();
        GUILayout.Space(30);

        if(GUILayout.Button("生成数据类"))
        {
            ExcelTool.CreateDataScript(excelFilesList);
            ShowNotification(new GUIContent("已生成数据类！"));
        }
        if(GUILayout.Button("读取数据并序列化"))
        {
            ExcelTool.AddDataToList(excelFilesList);
            ShowNotification(new GUIContent("已经数据读取到数据类并序列化！"));
        }
    }

    private void OnDestroy()
    {
        excelWindow = null;
    }
}