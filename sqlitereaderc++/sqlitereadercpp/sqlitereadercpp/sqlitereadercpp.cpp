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

