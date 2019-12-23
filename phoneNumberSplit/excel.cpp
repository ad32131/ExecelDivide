#include "stdafx.h"
#include "excel.h"
#include "stdio.h"
#include "comutil.h"
#include <atlconv.h>

BOOL excel::excelstart(HWND hDlg) {
	this->hDlg = hDlg;
	pXlApplication = NULL;

	hResult = OleInitialize(NULL);
	hResult = resultExceptionHandle(CLSIDFromProgID(OLESTR("Excel.Application"), &clsid));

	hResult = resultExceptionHandle(CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (LPVOID*)&pXlApplication));

	amiread = false;

	pXlApplication_create_id = pXlApplication;
	pXlApplication_read_id = pXlApplication;
	pXlApplication_close_id = pXlApplication;
	pXlApplication_quit_id = pXlApplication;
	pXlWorksheet = pXlApplication;
	pXlRangeCells = pXlApplication;
	return TRUE;
}

BOOL excel::excelcreatenewwork() {




	amiread = TRUE;
	lpszName = (OLECHAR*)OLESTR("Workbooks");
	setdispParams();
	hResult = resultExceptionHandle(pXlApplication_create_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

	setdispParams();
	hResult = resultExceptionHandle(pXlApplication_create_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL));

	pXlWorkbooks_createtmp = variant.pdispVal;

	
	lpszName = (OLECHAR*)OLESTR("Add");
	hResult = pXlWorkbooks_createtmp->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid);

	setdispParams();
	hResult = resultExceptionHandle( pXlWorkbooks_createtmp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL) );
	
	pXlWorkbook_createtmp = variant.pdispVal;


	pXlApplication_close_id = pXlWorkbooks_createtmp;
	return TRUE;
}

BOOL excel::excelSetActiveSheet(TCHAR* input) {
	lpszName = (OLECHAR*)OLESTR("ActiveSheet");


	dispParams.cArgs = 1;
	dispParams.cNamedArgs = 1;
	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];
	dispParams.rgvarg[0].vt = VT_BSTR;
	dispParams.rgvarg[0].bstrVal = SysAllocString( /*inputname*/L"Global Template");
	dispParams.rgdispidNamedArgs[0] = 0;
	VariantInit(&variant);
	hResult = pXlWorkbooks_createtmp->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid);

	hResult = resultExceptionHandle(pXlWorkbooks_createtmp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispParams, &variant, NULL, NULL));
	return TRUE;
}

BOOL excel::excelreadfile(OLECHAR* input) {
	//OLECHAR inputname[MAX_PATH] = OLESTR("C:\\testcc\\Mthetr_Avails.xlsx");
	lpszName = (OLECHAR*)OLESTR("Workbooks");
	setdispParams();
	hResult = resultExceptionHandle(pXlApplication_create_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

	setdispParams();
	hResult = resultExceptionHandle(pXlApplication_create_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL));

	pXlApplication_read_id = variant.pdispVal;

	setdispParams();
	lpszName = (OLECHAR*)OLESTR("Open");


	setdispParams();
	hResult = resultExceptionHandle(pXlApplication_read_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

	dispParams.cArgs = 1;
	dispParams.cNamedArgs = 1;
	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];
	dispParams.rgvarg[0].vt = VT_BSTR;
	dispParams.rgvarg[0].bstrVal = SysAllocString( /*inputname*/ input);
	dispParams.rgdispidNamedArgs[0] = 0;
	VariantInit(&variant);
	hResult = resultExceptionHandle(pXlApplication_read_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &variant, &excepInfo, NULL));

	SysFreeString(dispParams.rgvarg[0].bstrVal);
	delete[] dispParams.rgvarg;
	delete[] dispParams.rgdispidNamedArgs;

	pXlApplication_read_id = variant.pdispVal;


	return TRUE;
}

BOOL excel::excelDataSelect(OLECHAR* input) {
	uArgErr = 0;
	lpszName = (OLECHAR*)OLESTR("Cells");
	hResult = pXlWorksheet->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid);



	setdispParams();
	hResult = resultExceptionHandle(pXlWorksheet->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL));
	pXlRangeCells = variant.pdispVal;

	//OLESTR("$A$1")

	lpszName = (OLECHAR*)OLESTR("Range");
	hResult = resultExceptionHandle(pXlWorksheet->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

	dispParams.cArgs = 1;
	dispParams.cNamedArgs = 1;
	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];
	VariantInit(&dispParams.rgvarg[0]);
	dispParams.rgvarg[0].vt = VT_BSTR;
	//dispParams.rgvarg[0].bstrVal = SysAllocString(OLESTR("$A$1:$C$3")); range
	dispParams.rgvarg[0].bstrVal = SysAllocString(input);
	dispParams.rgdispidNamedArgs[0] = 0;
	VariantInit(&variant);
	hResult = resultExceptionHandle(pXlWorksheet->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, &excepInfo, &uArgErr));

	pXlRangeRange = variant.pdispVal;




	return TRUE;
}

double excel::excelDataRead(TCHAR* outputString, const int szint) {
	lpszName = (OLECHAR*)OLESTR("Value");
	hResult = resultExceptionHandle(pXlRangeRange->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

	dispParams.cArgs = 0;
	dispParams.cNamedArgs = 0;
	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];

	VariantInit(&variant);

	hResult = resultExceptionHandle(pXlRangeRange->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL));

	if (variant.vt == 0) {
		memset(outputString, 0, szint);
		return 0;
	}
	else if (variant.vt == 8) {//string
		//memcpy_s(outputString, MAX_PATH, variant.bstrVal, wcslen(variant.bstrVal)*2);
		memset(outputString, 0, szint);
		if (szint == 0) { return atoi((const char*)variant.pbVal); }
		if (variant.bstrVal != NULL) wcscpy_s(outputString, szint, variant.bstrVal);

		return 0;
	}
	else if (variant.vt == 5) {//double
		return variant.dblVal;
	}

	//MessageBox(hDlg, variant., L"Data", 0);

	delete[] dispParams.rgvarg;
	delete[] dispParams.rgdispidNamedArgs;


}

double excel::excelDataGetRead(TCHAR* outputString, const int szint, const TCHAR* condition_X, int condition_index) {

	TCHAR range_input[10];
	wsprintf(range_input, L"$%s$%d", condition_X, condition_index);
	excelDataSelect(range_input);
	return excelDataRead(outputString, szint);
}


BOOL excel::excelDataWrite(OLECHAR* input) {
	lpszName = (OLECHAR*)OLESTR("Value");
	hResult = resultExceptionHandle(pXlRangeRange->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

	dispParams.cArgs = 1;
	dispParams.cNamedArgs = 1;
	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];

	dispParams.rgvarg[0].vt = VT_BSTR;
	dispParams.rgvarg[0].bstrVal = SysAllocString(input);
	dispParams.rgvarg[0].cVal;

	dispParams.rgdispidNamedArgs[0] = DISPID_PROPERTYPUT;

	VariantInit(&variant);

	hResult = resultExceptionHandle(pXlRangeRange->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &dispParams, &variant, NULL, NULL));


	delete[] dispParams.rgvarg;
	delete[] dispParams.rgdispidNamedArgs;

	return TRUE;
}

BOOL excel::excelWriteColor(OLECHAR* input, unsigned int color) {

	//  Interior  Range 
	lpszName = (OLECHAR*)OLESTR("Interior");
	VariantInit(&variant);
	hResult = pXlRangeRange->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid);

	dispParams = { NULL, NULL, 0, 0 };

	hResult = pXlRangeRange->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, &excepInfo, &uArgErr);

	pXlRangeColor = variant.pdispVal;

	// ColorInterior 
	DISPID dispIDColor;
	lpszName = (OLECHAR*)OLESTR("Color");
	hResult = pXlRangeColor->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispIDColor);


	dispParams.cArgs = 1;
	dispParams.cNamedArgs = 1;

	dispParams.rgvarg = new VARIANTARG[1];
	dispParams.rgdispidNamedArgs = new DISPID[1];
	dispParams.rgvarg[0].vt = VT_I4;
	dispParams.rgvarg[0].lVal = color;
	dispParams.rgdispidNamedArgs[0] = DISPID_PROPERTYPUT;
	VariantInit(&variant);

	hResult = pXlRangeColor->Invoke(dispIDColor, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dispParams, &variant, &excepInfo, &uArgErr);

	return TRUE;
}


BOOL excel::excelsaveas(int rw_type, OLECHAR* input) {

	if (rw_type == 1) {

		lpszName = (OLECHAR*)OLESTR("SaveAs");
		hResult = resultExceptionHandle(pXlWorkbook_createtmp->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

		//pXlWorkbook_createtmp
		//pXlApplication_read_id
		USES_CONVERSION;
		remove(OLE2A(SysAllocString(input)));
		dispParams.cArgs = 1;
		dispParams.cNamedArgs = 1;
		dispParams.rgvarg = new VARIANTARG[1];
		dispParams.rgdispidNamedArgs = new DISPID[1];
		dispParams.rgvarg[0].vt = VT_BSTR;
		dispParams.rgvarg[0].bstrVal = SysAllocString(input);
		dispParams.rgdispidNamedArgs[0] = (DISPID)0;


		VariantInit(&variant);
		hResult = resultExceptionHandle(pXlWorkbook_createtmp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &variant, NULL, NULL));
		SysFreeString(dispParams.rgvarg[0].bstrVal);
		delete[] dispParams.rgvarg;
		delete[] dispParams.rgdispidNamedArgs;

		return TRUE;
	}
	else if (rw_type == 2) {
		lpszName = (OLECHAR*)OLESTR("SaveAs");
		hResult = resultExceptionHandle(pXlApplication_read_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

		USES_CONVERSION;
		remove(OLE2A(SysAllocString(input)));
		dispParams.cArgs = 1;
		dispParams.cNamedArgs = 1;
		dispParams.rgvarg = new VARIANTARG[1];
		dispParams.rgdispidNamedArgs = new DISPID[1];
		dispParams.rgvarg[0].vt = VT_BSTR;
		dispParams.rgvarg[0].bstrVal = SysAllocString(input);
		dispParams.rgdispidNamedArgs[0] = (DISPID)0;


		VariantInit(&variant);
		hResult = resultExceptionHandle(pXlApplication_read_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &variant, NULL, NULL));
		SysFreeString(dispParams.rgvarg[0].bstrVal);
		delete[] dispParams.rgvarg;
		delete[] dispParams.rgdispidNamedArgs;

		return TRUE;
	}
}

BOOL excel::excelsave(int rw_type) {
	if (rw_type == 1) {
		lpszName = (OLECHAR*)OLESTR("Save");
		hResult = resultExceptionHandle(pXlWorkbook_createtmp->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));



		dispParams.cArgs = 0;
		dispParams.cNamedArgs = 0;
		dispParams.rgvarg = NULL;
		dispParams.rgdispidNamedArgs = NULL;
		VariantInit(&variant);
		hResult = resultExceptionHandle(pXlWorkbook_createtmp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &variant, NULL, NULL));

		return TRUE;
	}
	else if (rw_type == 2)
	{
		lpszName = (OLECHAR*)OLESTR("Save");
		hResult = resultExceptionHandle(pXlApplication_read_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));



		dispParams.cArgs = 0;
		dispParams.cNamedArgs = 0;
		dispParams.rgvarg = NULL;
		dispParams.rgdispidNamedArgs = NULL;
		VariantInit(&variant);
		hResult = resultExceptionHandle(pXlApplication_read_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &variant, NULL, NULL));

		return TRUE;
	}
}

BOOL excel::excelclosefile() {
	if (amiread == FALSE) return FALSE;
	lpszName = (OLECHAR*)OLESTR("Close");
	hResult = resultExceptionHandle(pXlApplication_close_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

	setdispParams();
	hResult = resultExceptionHandle(pXlApplication_close_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &variant, NULL, NULL));

	pXlApplication_close_id = variant.pdispVal;

	return TRUE;
}

BOOL excel::excelquit() {
	if (amiread == FALSE) return FALSE;
	hResult = OleInitialize(NULL);
	hResult = CLSIDFromProgID(OLESTR("Excel.Application"), &clsid);

	lpszName = (OLECHAR*)OLESTR("Quit");
	hResult = resultExceptionHandle(pXlApplication_quit_id->GetIDsOfNames(IID_NULL, &lpszName, 1, LOCALE_USER_DEFAULT, &dispid));

	setdispParams();
	hResult = resultExceptionHandle(pXlApplication_quit_id->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &variant, NULL, NULL));


	dispatchUnInit(pXlApplication);
	dispatchUnInit(pXlApplication_create_id);
	dispatchUnInit(pXlApplication_read_id);
	dispatchUnInit(pXlApplication_close_id);
	dispatchUnInit(pXlApplication_quit_id);

	dispatchUnInit(pXlWorkbooks_createtmp);
	dispatchUnInit(pXlWorkbook_createtmp);
	dispatchUnInit(pXlWorksheets);
	dispatchUnInit(pXlWorksheet);
	dispatchUnInit(pXlRangeCells);

	dispatchUnInit(pXlRangeCell);
	dispatchUnInit(pXlRangeRange);
	dispatchUnInit(pXlRangeColor);
	OleUninitialize();
	return TRUE;
}

BOOL excel::dispatchUnInit(IDispatch* dis) {
	if (((short)dis) <= 0) return FALSE;
	else if (dis && (((short)dis) != -1) && ((short)dis != 0x2)) dis->Release();
	return TRUE;

};

BOOL excel::setdispParams() {
	dispParams.cArgs = 0;
	dispParams.cNamedArgs = 0;
	dispParams.rgvarg = (VARIANTARG*)NULL;
	dispParams.rgdispidNamedArgs = (DISPID*)NULL;
	VariantInit(&variant);

	return TRUE;
}

HRESULT excel::resultExceptionHandle(HRESULT hResult) {
	if (hResult == 0) {
		return hResult;
	}
	else if (hResult == -2147352573) {
		return hResult;
	}
	else {

		TCHAR error[MAX_PATH];
		wsprintf(error, L"0x%p CLASS %s ERROR!", hResult, lpszName);
		//MessageBox(this->hDlg, error, L"error", 0);
		/*
		excelquit();
		exit(0);
		*/
		return hResult;
	}
};

BOOL excel::excelDataGreenSet(OLECHAR* write_contents, const TCHAR* condition_X, int condition_index) {
	TCHAR range_input[10];
	wsprintf(range_input, L"$%s$%d", condition_X, condition_index);
	excelDataSelect(range_input);

	//excel1.excelDataRead(getData);

	excelWriteColor(range_input, 65280);
	excelDataWrite(write_contents);
	return TRUE;
}

BOOL excel::excelDataSet(OLECHAR* write_contents, const TCHAR* condition_X, int condition_index) {
	TCHAR range_input[10];
	wsprintf(range_input, L"$%s$%d", condition_X, condition_index);
	excelDataSelect(range_input);

	excelDataWrite(write_contents);
	return TRUE;
}

BOOL excel::excelDataYellowSet(OLECHAR* write_contents, const TCHAR* condition_X, int condition_index) {
	TCHAR range_input[10];
	wsprintf(range_input, L"$%s$%d", condition_X, condition_index);
	excelDataSelect(range_input);

	//excel1.excelDataRead(getData);

	excelWriteColor(range_input, 0x00FFFF);
	excelDataWrite(write_contents);
	return TRUE;
}


excel::~excel() {

	dispatchUnInit(pXlApplication);
	dispatchUnInit(pXlApplication_create_id);
	dispatchUnInit(pXlApplication_read_id);
	dispatchUnInit(pXlApplication_close_id);
	dispatchUnInit(pXlApplication_quit_id);

	dispatchUnInit(pXlWorkbooks_createtmp);
	dispatchUnInit(pXlWorkbook_createtmp);
	dispatchUnInit(pXlWorksheets);
	dispatchUnInit(pXlWorksheet);
	dispatchUnInit(pXlRangeCells);

	dispatchUnInit(pXlRangeCell);
	dispatchUnInit(pXlRangeRange);
	dispatchUnInit(pXlRangeColor);

	OleUninitialize();
}