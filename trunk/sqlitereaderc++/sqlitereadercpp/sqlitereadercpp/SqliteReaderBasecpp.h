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
#pragma once

#include "sqlite3.h"

typedef struct _Text_Value
{
	int nBufLen;
	const unsigned char* pBuf;
}t_Text_Value;

typedef union
{
	int nValue;
	t_Text_Value txtValue;
	__int64 n64Value;
	double dbValue;
}unionValue;


class CSqliteReaderBasecpp
{
public:
	CSqliteReaderBasecpp(void);
	~CSqliteReaderBasecpp(void);
public:
	bool OpenDB(char* pDBFileName);
	void Clear();
	bool PrepareSql(char* pSql);	
	bool Fetch();
	bool GetOneRowData(int nFieldCount, unionValue values[], int nFieldType[]);	
private:
	sqlite3* m_pDB;
	sqlite3_stmt* m_pStmt;
};

