#pragma once
#include <afxdisp.h>

class CMSWord
{
public:
	CMSWord();
	~CMSWord();

	bool CreateWordInstance();
	bool CreateNewDocument();
	bool AddTitle(const CString& title);
	bool AddParagraph(const CString& text);
	bool CreateTable(int rows, int cols);
	bool SetTableCell(int row, int col, const CString& text);
	bool SaveAs(const CString& filePath);
	void CloseWord();
	LPDISPATCH m_pDocument;
	LPDISPATCH m_pSelection;

private:
	LPDISPATCH m_pWordApp;
	LPDISPATCH m_pDocuments;
	
	LPDISPATCH m_pTable;
	bool m_bInit;
};
