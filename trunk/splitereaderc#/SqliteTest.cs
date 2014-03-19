/*
The MIT License (MIT)

Copyright (c) <2013-2020> <wenshengming zhujiangping>

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

.
 */

using UnityEngine;
using System.Collections;
using System.Data.Common;
using Mono.Data.Sqlite;
using AssemblyCSharp;

public class SqliteTest : MonoBehaviour 
{
	private SqliteConnection m_sqliteConnection;
	private bool m_bPrinted;
	// Use this for initialization
	SqliteTest()
	{
		m_bPrinted = false;
	}
	void Start () 
	{
//		string strCon = @"data source=f:\\wsm\\SampleDB";
//		try
//		{
//			m_sqliteConnection = new SqliteConnection(strCon);
//			m_sqliteConnection.Open();
//			Debug.Log("open sucess");
//		}
//		catch(SqliteException e)
//		{
//			Debug.Log(e.ToString());
//		}


//		SQliteReader sqReader = new SQliteReader();
//		if (!sqReader.OpenDB("f:\\wsm\\SampleDB"))
//		{
//			return ;
//		}
//		Field_Info[] fieldInfo = new Field_Info[3];
//		fieldInfo[0].eFieldType = E_Sqlite_Field_Type.E_Sqlite_Field_Type_Int;
//		fieldInfo[0].strFieldName = "RecordOrder";
//		fieldInfo[1].eFieldType = E_Sqlite_Field_Type.E_Sqlite_Field_Type_Int;
//		fieldInfo[1].strFieldName = "classid";
//		fieldInfo[2].eFieldType = E_Sqlite_Field_Type.E_Sqlite_Field_Type_Varchar;
//		fieldInfo[2].strFieldName = "classname";
//
//		if (!sqReader.OpenTable("classdefine", fieldInfo))
//		{
//			return ;
//		}
//		while (sqReader.ReadNext())
//		{
//			int nFieldOrder = 0;
//			int nValue = 0;
//			string strValue = "";
//			sqReader.GetInt(0, ref nFieldOrder);
//			sqReader.GetInt(1, ref nValue);
//			sqReader.GetString(2, ref strValue);
//		}
	}
	
	// Update is called once per frame
	void Update () 
	{
//		if (!m_bPrinted)
//		{
//			SqliteCommand sqcommand = m_sqliteConnection.CreateCommand();
//			sqcommand.CommandText = "select * from classdefine";
//			SqliteDataReader sqReader = sqcommand.ExecuteReader();
//			while (sqReader.Read())
//			{
//				Debug.Log(sqReader.GetValue(0).ToString());
//				Debug.Log(sqReader.GetValue(1).ToString());
//				Debug.Log(sqReader.GetValue(2).ToString());
//			}
//			sqReader.Close();
//			m_bPrinted = true;
//		}
	}
}
