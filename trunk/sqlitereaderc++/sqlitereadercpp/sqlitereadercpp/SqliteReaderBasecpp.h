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

