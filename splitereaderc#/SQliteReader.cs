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

using System;
using System.Data.Common;
using Mono.Data.Sqlite;
using UnityEngine;

namespace AssemblyCSharp
{
	public enum E_Sqlite_Field_Type
	{
		E_Sqlite_Field_Type_Invalid,
		E_Sqlite_Field_Type_Int,
		E_Sqlite_Field_Type_Varchar,
		E_Sqlite_Field_Type_Float,
		E_Sqlite_Field_Type_Short,
		E_Sqlite_Field_Type_Byte,
		E_Sqlite_Field_Type_Int64,
	};
	public struct Field_Info
	{
		public E_Sqlite_Field_Type eFieldType;
		public string strFieldName;
	};
	public class SQliteReader
	{
		private SqliteConnection m_SqliteConnection;
		private Field_Info[] m_ArrayFieldInfo;
		private SqliteDataReader m_sqDataReader;

		public SQliteReader ()
		{
			m_SqliteConnection = null;
			m_sqDataReader = null;
			m_ArrayFieldInfo = null;
		}

		public bool OpenDB(string strDBFileName)
		{
			if (strDBFileName == null)
			{
				return false;
			}
			if (strDBFileName == "")
			{
				return false;
			}
			string strConString = @"data source=" + strDBFileName;

			m_SqliteConnection = new SqliteConnection(strConString);
			try
			{
				m_SqliteConnection.Open();
				Debug.Log(strConString + " Open Success!!");
				return true;
			}
			catch(Exception E)
			{
				Debug.LogError(strConString + " Open Failed," + E.ToString());
				m_SqliteConnection = null;
				return false;
			}
		}

		public bool OpenTable(string strTableName, Field_Info[] fieldInfo)
		{
			if (m_SqliteConnection == null)
			{
				return false;
			}
			if (strTableName == null)
			{
				return false;
			}
			if (strTableName == "")
			{
				return false;
			}
			if (fieldInfo == null)
			{
				return false;
			}
			if (fieldInfo.Length == 0)
			{
				return false;
			}
			m_ArrayFieldInfo = fieldInfo;

			if (m_sqDataReader != null)
			{
				m_sqDataReader.Close();
				m_sqDataReader = null;
			}

			SqliteCommand sqCommand = m_SqliteConnection.CreateCommand();
			sqCommand.CommandText = "select ";
			for (int i = 0; i < this.m_ArrayFieldInfo.Length; ++i)
			{
				if (i == 0)
				{
					sqCommand.CommandText += this.m_ArrayFieldInfo[i].strFieldName;
				}
				else
				{
					sqCommand.CommandText += "," + this.m_ArrayFieldInfo[i].strFieldName;
				}
			}
			sqCommand.CommandText += " from " + strTableName;

			try
			{
				m_sqDataReader = sqCommand.ExecuteReader();
				return true;
			}
			catch(Exception E)
			{
				Debug.Log("Table Open Failed!!!" + E.ToString());
				return false;
			}
		}
		public bool ReadNext()
		{
			if (this.m_sqDataReader == null)
			{
				return false;
			}
			return m_sqDataReader.Read();
		}
		private bool CheckFieldIndexValidAndTypeValid(int nFieldIndex, E_Sqlite_Field_Type eType)
		{
			if (m_sqDataReader == null)
			{
				return false;
			}

			if (nFieldIndex < 0 || nFieldIndex > m_sqDataReader.FieldCount || nFieldIndex > this.m_ArrayFieldInfo.Length)
			{
				return false;
			}
			if (m_ArrayFieldInfo[nFieldIndex].eFieldType != eType)
			{
				return false;
			}
			return true;
		}
		public bool GetByte(int nFieldIndex, ref byte byValue)
		{
			if (!CheckFieldIndexValidAndTypeValid(nFieldIndex, E_Sqlite_Field_Type.E_Sqlite_Field_Type_Byte))
			{
				return false;
			}
			byValue = m_sqDataReader.GetByte(nFieldIndex);
			return true;
		}
		public bool GetShort(int nFieldIndex, ref short sValue)
		{
			if (!CheckFieldIndexValidAndTypeValid(nFieldIndex, E_Sqlite_Field_Type.E_Sqlite_Field_Type_Short))
			{
				return false;
			}
			sValue = m_sqDataReader.GetInt16(nFieldIndex);
			return true;
		}
		public bool GetInt(int nFieldIndex, ref int nValue)
		{
			if (!CheckFieldIndexValidAndTypeValid(nFieldIndex, E_Sqlite_Field_Type.E_Sqlite_Field_Type_Int))
			{
				return false;
			}
			nValue = m_sqDataReader.GetInt32(nFieldIndex);
			return true;
		}
		public bool GetInt64(int nFieldIndex, ref Int64 n64Value)
		{
			if (!CheckFieldIndexValidAndTypeValid(nFieldIndex, E_Sqlite_Field_Type.E_Sqlite_Field_Type_Int64))
			{
				return false;
			}
			n64Value = m_sqDataReader.GetInt64(nFieldIndex);
			return true;
		}
		public bool GetFloat(int nFieldIndex, ref float fValue)
		{
			if (!CheckFieldIndexValidAndTypeValid(nFieldIndex, E_Sqlite_Field_Type.E_Sqlite_Field_Type_Float))
			{
				return false;
			}
			fValue = m_sqDataReader.GetFloat(nFieldIndex);
			return true;
		}
		public bool GetString(int nFieldIndex, ref string strValue)
		{
			if (!CheckFieldIndexValidAndTypeValid(nFieldIndex, E_Sqlite_Field_Type.E_Sqlite_Field_Type_Varchar))
			{
				return false;
			}
			strValue = m_sqDataReader.GetString(nFieldIndex);
			return true;
		}
	}
}

