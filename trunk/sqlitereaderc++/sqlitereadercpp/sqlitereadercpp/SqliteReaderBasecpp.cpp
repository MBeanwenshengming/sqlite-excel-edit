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
#include "stdafx.h"
#include "SqliteReaderBasecpp.h"
#include <string>
using namespace std;
#include "windows.h"

CSqliteReaderBasecpp::CSqliteReaderBasecpp(void)
{
	m_pDB = NULL;
	m_pStmt = NULL;
}


CSqliteReaderBasecpp::~CSqliteReaderBasecpp(void)
{
	if (m_pStmt != NULL)
	{
		sqlite3_finalize(m_pStmt);
		m_pStmt = NULL;
	}
	if (m_pDB != NULL)
	{
		sqlite3_close(m_pDB);
		m_pDB = NULL;
	}
}


bool CSqliteReaderBasecpp::OpenDB(char* pDBFileName)
{
	if (pDBFileName == NULL)
	{
		return false;
	}
	int nResult = sqlite3_open(pDBFileName, &m_pDB);
	if (nResult != SQLITE_OK)
	{
		return false;
	}
	return true;
}


void CSqliteReaderBasecpp::Clear()
{
	if (m_pDB == NULL)
	{
		return ;
	}
	if (m_pStmt != NULL)
	{
		sqlite3_finalize(m_pStmt);
		m_pStmt = NULL;
	}
}
bool CSqliteReaderBasecpp::PrepareSql(char* pSql)
{
	if (pSql == NULL)
	{
		return false;
	}
	Clear();

	const char* pszTail = NULL;
	if (sqlite3_prepare_v2(m_pDB, pSql, strlen(pSql), &m_pStmt,  &pszTail) != SQLITE_OK)
	{
		return false;
	}
	return true;
}
bool CSqliteReaderBasecpp::Fetch()
{
	if (m_pDB == NULL)
	{
		return false;
	}
	if (m_pStmt == NULL)
	{
		return false;
	}
	if (sqlite3_step(m_pStmt) != SQLITE_ROW)
	{
		return false;
	}
	return true;
}


//	http://www.cnblogs.com/stephen-liu74/archive/2012/01/18/2325258.html
bool CSqliteReaderBasecpp::GetOneRowData(int nFieldCount, unionValue values[], int nFieldType[])
{
	if (m_pStmt == NULL)
	{
		return false;
	}
	int nColumn = sqlite3_column_count(m_pStmt);
	if (nColumn != nFieldCount)
	{
		return false;
	}
	WCHAR   * pWstr = NULL;
	for (int i = 0; i < nFieldCount; ++i)
	{
		nFieldType[i] = sqlite3_column_type(m_pStmt, i);
		switch(nFieldType[i])
		{
		case SQLITE_INTEGER:
			{
				values[i].nValue = sqlite3_column_int(m_pStmt, i);
			}
			break;
		case SQLITE_FLOAT:
			{
				values[i].dbValue = sqlite3_column_double(m_pStmt, i);
			}
			break;
		case SQLITE_BLOB:
			{
				values[i].txtValue.pBuf		= (const unsigned char*)sqlite3_column_blob(m_pStmt, i);
				values[i].txtValue.nBufLen = sqlite3_column_bytes(m_pStmt, i);						
			}
			break;
		case SQLITE_NULL:
			{

			}
			break;
		case SQLITE_TEXT:
			{
				values[i].txtValue.pBuf		= sqlite3_column_text(m_pStmt, i);
				values[i].txtValue.nBufLen = sqlite3_column_bytes(m_pStmt, i);			

				//	数据库存储的为utf8字符串,  test  utf8 => multibyte
				//int     nLen = ::MultiByteToWideChar (CP_UTF8, 0, (const char*)values[i].txtValue.pBuf, -1, NULL, 0) ;
				//pWstr = new WCHAR[nLen+1] ;
				//ZeroMemory (pWstr, sizeof(WCHAR) * (nLen+1)) ;
				//::MultiByteToWideChar (CP_UTF8, 0, (const char*)values[i].txtValue.pBuf, -1, pWstr, nLen) ;

				//delete[] pWstr;
			}
			break;
		}

		//string stype = sqlite3_column_decltype(m_pStmt, i);
		//stype = strlwr((char*)stype.c_str());  

		//if (stype.find("int") != string::npos) 
		//{
		//	values[i].nValue = sqlite3_column_int(m_pStmt, i);
		//}
		//else if (stype.find("char") != string::npos  || stype.find("text") != string::npos) 
		//{			
		//	values[i].txtValue.nBufLen = sqlite3_column_bytes(m_pStmt, i);
		//	values[i].txtValue.pBuf = sqlite3_column_text(m_pStmt, i);
		//}
		//else if (stype.find("real") != string::npos || stype.find("floa") != string::npos || stype.find("doub") != string::npos )
		//{
		//	values[i].dbValue = sqlite3_column_double(m_pStmt, i);			
		//}
		//else if (stype.find("numeric") != string::npos || stype.find("boolean") != string::npos || stype.find("date") != string::npos || stype.find("datetime") != string::npos)
		//{
		//	values[i].n64Value = sqlite3_column_int64(m_pStmt, i);		
		//}
		//else
		//{
		//	continue;
		//}
	}
	return true;
}