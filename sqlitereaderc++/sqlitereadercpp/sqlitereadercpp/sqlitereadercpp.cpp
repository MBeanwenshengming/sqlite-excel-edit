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

// sqlitereadercpp.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include "sqlite3.h"
#include "SqliteReaderBasecpp.h"

int main(int argc, char* argv[])
{
	CSqliteReaderBasecpp sqReader;
	if (!sqReader.OpenDB("E:\\sqliteDB\\SDB12"))
	{
		return -1;
	}
	if (!sqReader.PrepareSql("select * from tabledefine"))
	{
		return -1;
	}
	int nFieldCount = 9;
	unionValue unionValue[9];
	int nFieldType[9];
	while (sqReader.Fetch())
	{		
		sqReader.GetOneRowData(nFieldCount, unionValue, nFieldType);		
	}
	return 0;
}

