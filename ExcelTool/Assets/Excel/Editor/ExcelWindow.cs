//
// Excel����UI����
//
using UnityEngine;
using UnityEditor;
using System.Collections.Generic;
using System.IO;

public class ExcelWindow : EditorWindow
{
    private static ExcelWindow excelWindow;
    public List<FileData> excelFilesList = new List<FileData>();

    [MenuItem("Tools/Excel����", false, -10)]
    static void ShowWindow()
    {
        excelWindow = EditorWindow.GetWindow<ExcelWindow>();
        excelWindow.titleContent = new GUIContent("Excel����");
        excelWindow.position = new Rect(200, 300, 800, 400);
    }

    private void OnGUI()
    {
        GUILayout.Label("Excel����", "WarningOverlay");
        if(GUILayout.Button("��ȡ·���µ�Excel�ļ�"))
        {
            ExcelTool.GetAllExcelFilesPath(excelFilesList);
        }
        GUILayout.BeginVertical("box");
        foreach(FileData fileData in excelFilesList)
        {
            GUILayout.BeginHorizontal();
            fileData.isSelect = EditorGUILayout.Toggle(fileData.isSelect, GUILayout.Width(20));
            GUILayout.Label("�ļ�����" + fileData.excelFileName);
            GUILayout.Space(30);
            GUILayout.Label("·��" + fileData.excelFilePath);
            GUILayout.EndHorizontal();
            GUILayout.Space(5);
        }
        GUILayout.EndVertical();
        GUILayout.Space(30);

        if(GUILayout.Button("����������"))
        {
            ExcelTool.CreateDataScript(excelFilesList);
            ShowNotification(new GUIContent("�����������࣡"));
        }
        if(GUILayout.Button("��ȡ���ݲ����л�"))
        {
            ExcelTool.AddDataToList(excelFilesList);
            ShowNotification(new GUIContent("�Ѿ����ݶ�ȡ�������ಢ���л���"));
        }
    }

    private void OnDestroy()
    {
        excelWindow = null;
    }
}