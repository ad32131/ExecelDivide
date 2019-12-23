#pragma once
#include <OAIdl.h>

class excel {
private:
public:
	HWND hDlg;
	HRESULT hResult;
	CLSID clsid;
	UINT uArgErr;
	BOOL amiread;

	//action
	IDispatch* pXlApplication = NULL;
	IDispatch* pXlApplication_create_id = NULL;
	IDispatch* pXlApplication_read_id = NULL;
	IDispatch* pXlApplication_close_id = NULL;
	IDispatch* pXlApplication_quit_id = NULL;

	//create
	IDispatch* pXlWorkbooks_createtmp = NULL;
	IDispatch* pXlWorkbook_createtmp = NULL;

	//readData
	IDispatch* pXlWorksheets = NULL;
	IDispatch* pXlWorksheet = NULL;
	IDispatch* pXlRangeCells = NULL;
	IDispatch* pXlRangeCell = NULL;
	IDispatch* pXlRangeRange = NULL;
	IDispatch* pXlRangeColor = NULL;


	EXCEPINFO excepInfo;
	DISPID dispid;
	OLECHAR* lpszName;
	DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
	VARIANT variant;


	BOOL excelstart(HWND hDlg);
	BOOL excelcreatenewwork();
	BOOL excelSetActiveSheet(OLECHAR* input);
	BOOL excelreadfile(OLECHAR* input);

	BOOL excelsave(int rw_type);
	BOOL excelclosefile();
	BOOL excelquit();

	BOOL dispatchUnInit(IDispatch* dis);
	HRESULT resultExceptionHandle(HRESULT hResult);
	BOOL setdispParams();

	BOOL excelDataSelect(OLECHAR* input);
	double excelDataRead(TCHAR* outputString, const int szint);
	BOOL excelDataWrite(OLECHAR* input);
	BOOL excelWriteColor(OLECHAR* input, unsigned int color);

	double excelDataGetRead(TCHAR* outputString, const int szint, const TCHAR* condition_X, int condition_index);
	BOOL excelDataGreenSet(OLECHAR* write_contents, const TCHAR* condition_X, int condition_index);
	BOOL excelDataYellowSet(OLECHAR* write_contents, const TCHAR* condition_X, int condition_index);
	BOOL excelDataSet(OLECHAR* write_contents, const TCHAR* condition_X, int condition_index);

	BOOL excelsaveas(int rw_type, OLECHAR* input);

	~excel();
};