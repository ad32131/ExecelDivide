// phoneNumberSplit.cpp : 응용 프로그램에 대한 진입점을 정의합니다.
//

#include "stdafx.h"
#include "phoneNumberSplit.h"
#include "excel.h"
#include "resource.h"
#include <CommCtrl.h> 
#include <shellapi.h>
#include <shlobj_core.h>
#include <Commdlg.h>
#include <TlHelp32.h>
#include <Shlwapi.h>
#include <stdlib.h>
#include <stdio.h>


#pragma comment(lib, "Shlwapi.lib")

#define MAX_LOADSTRING 100

// 전역 변수:
HINSTANCE hInst;                                // 현재 인스턴스입니다.
WCHAR szTitle[MAX_LOADSTRING];                  // 제목 표시줄 텍스트입니다.
WCHAR szWindowClass[MAX_LOADSTRING];            // 기본 창 클래스 이름입니다.
int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM /*lParam*/, LPARAM lpData);

BOOL processAllKill(HWND hDlg, const WCHAR* szProcessName);
BOOL CALLBACK DialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR    lpCmdLine, _In_ int       nCmdShow);
unsigned char randomchr();
void OnDrawItem(HWND ah_dlg, UINT a_ctrl_id, DRAWITEMSTRUCT *ap_dis, unsigned int IDC_INT);
LRESULT getListColor(WPARAM wParam, HBRUSH hBrush);

BOOL CALLBACK DialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
	static BROWSEINFO browse_OutPutFilePath;
	
	static HWND IDC_GETFILELIST_H;
	static unsigned int dragFileTimes = 0;
	static HWND IDC_SETFILEPATH_H;
	static HWND IDC_OUTFILEPATH_H;
	static HWND IDC_PROGRESS1_H;
	static HWND IDC_STATIC_H;
	static HWND IDC_COUNT_H;
	static HWND IDC_LINECOUNT_H;
	static HWND IDOK_H;
	static HWND IDC_SET_H;

	LPITEMIDLIST pidl;
	static OLECHAR outPutFilePath[MAX_PATH];
	static OLECHAR inReadFilePath[MAX_PATH];
	static OLECHAR inReadFilePathTemp[MAX_PATH];
	static OLECHAR intLineDiv[MAX_PATH];
	static OLECHAR readText[MAX_PATH];
	static OLECHAR randomText[30];
	TCHAR filenamefromdrag[MAX_PATH];
	static char readTextTemp[MAX_PATH * 2];

	static HDROP hDrop = 0;
	static FILE *fp;
	static unsigned int line = 0;
	static unsigned int linediv = 0;
	static unsigned int lineindex = 0;
	static char c;
	static excel outExcelFile;
	static HBRUSH hBrush;

	switch (message) {
	case WM_INITDIALOG:
		IDC_GETFILELIST_H = GetDlgItem(hDlg, IDC_GETFILELIST);
		IDC_SETFILEPATH_H = GetDlgItem(hDlg, IDC_SETFILEPATH);
		IDC_OUTFILEPATH_H = GetDlgItem(hDlg, IDC_OUTFILEPATH);
		IDC_PROGRESS1_H = GetDlgItem(hDlg, IDC_PROGRESS1);
		IDC_STATIC_H = GetDlgItem(hDlg, IDC_STATIC);
		IDC_COUNT_H = GetDlgItem(hDlg, IDC_COUNT);
		IDC_LINECOUNT_H = GetDlgItem(hDlg, IDC_LINECOUNT);
		IDOK_H = GetDlgItem(hDlg, IDOK);
		IDC_SET_H = GetDlgItem(hDlg, IDC_SET);

		hBrush = CreateSolidBrush(RGB(0, 0,0));

		memset(outPutFilePath, 0, MAX_PATH);
		SHGetSpecialFolderPath(NULL, outPutFilePath, CSIDL_DESKTOPDIRECTORY, 0);

		SetWindowText(IDC_OUTFILEPATH_H, outPutFilePath);
		SetWindowText(IDC_COUNT_H, L"100");
		DragAcceptFiles(hDlg, TRUE);

		processAllKill(hDlg, L"EXCEL.EXE");
		

		break;
	case WM_DRAWITEM:
		OnDrawItem(IDC_SETFILEPATH_H, wParam, (DRAWITEMSTRUCT *)lParam, IDC_SETFILEPATH);
		OnDrawItem(IDOK_H, wParam, (DRAWITEMSTRUCT *)lParam, IDOK);
		OnDrawItem(IDC_SET_H, wParam, (DRAWITEMSTRUCT *)lParam, IDC_SET);
		break;
	case WM_COMMAND:
		switch (LOWORD(wParam)) {
		case IDC_SETFILEPATH:
			memset(&browse_OutPutFilePath, 0, sizeof(BROWSEINFO));
			memset(outPutFilePath, 0, MAX_PATH);
			browse_OutPutFilePath.hwndOwner = NULL;
			browse_OutPutFilePath.pidlRoot = NULL;
			browse_OutPutFilePath.pszDisplayName = outPutFilePath;
			browse_OutPutFilePath.lpszTitle = L"PATH SET";
			browse_OutPutFilePath.ulFlags = BIF_RETURNONLYFSDIRS | BIF_STATUSTEXT | BIF_VALIDATE;
			browse_OutPutFilePath.lpfn = BrowseCallbackProc;
			browse_OutPutFilePath.lParam = (LPARAM)outPutFilePath;
			browse_OutPutFilePath.ulFlags = BIF_RETURNONLYFSDIRS;

			pidl = SHBrowseForFolder(&browse_OutPutFilePath);
			if (pidl != NULL) {
				SHGetPathFromIDList(pidl, outPutFilePath);
				SetWindowText(IDC_OUTFILEPATH_H, outPutFilePath);
			}
			break;
		case IDC_GETFILELIST:
			switch (HIWORD(wParam)) {
				case LBN_DBLCLK:
					int copy_index = SendMessage(IDC_GETFILELIST_H, LB_GETCURSEL, 0, 0);
					SendMessage(IDC_GETFILELIST_H, LB_DELETESTRING, copy_index, 0);
					break;
			}
			break;
		case IDC_SET:
			line = 0;
			memset(inReadFilePath, 0, MAX_PATH);
			SendMessage(IDC_GETFILELIST_H, LB_GETTEXT, 0, (LPARAM)inReadFilePath);
			_wfopen_s(&fp, inReadFilePath, L"r");

			while ((c = fgetc(fp)) != EOF)
				if (c == '\n') line++;
			
			fclose(fp);

			line++;
			memset(readText, 0, MAX_PATH);
			memset(readTextTemp, 0, MAX_PATH);
			wsprintf(readText, L"%d", line);
			SetWindowText(IDC_LINECOUNT_H, readText);
			SetWindowText(IDC_COUNT_H, readText);
			break;
		case IDOK:
			GetWindowText(IDC_COUNT_H, intLineDiv, MAX_PATH);
			linediv = _wtoi(intLineDiv);

			SendMessage(IDC_PROGRESS1_H, PBM_SETRANGE, 0, MAKELPARAM(0, line));
			SendMessage(IDC_PROGRESS1_H, PBM_SETPOS, (WPARAM)0, (LPARAM)NULL);

			memset(inReadFilePath, 0, MAX_PATH);
			memset(inReadFilePathTemp, 0, MAX_PATH);
			SendMessage(IDC_GETFILELIST_H, LB_GETTEXT, 0, (LPARAM)inReadFilePath);

			_wfopen_s(&fp, inReadFilePath, L"r");
			
			//IDC_PROGRESS1_H
			
			lineindex = 0;
			
			for( int write_int=0; write_int < (int)(line / linediv) + 1; write_int++){
				outExcelFile.excelstart(hDlg);
				outExcelFile.excelcreatenewwork();
				wsprintf(inReadFilePathTemp, L"%s_%d.xls", inReadFilePath, write_int+1);

				for (int read_count = 0; read_count < linediv; read_count++, SendMessage(IDC_PROGRESS1_H, PBM_SETPOS, (WPARAM)++lineindex, (LPARAM)NULL)) {
				
					memset(readText, 0, MAX_PATH);
					fgetws(readText, MAX_PATH, fp);

					if ( !wcslen(readText)) {
						continue;
					}
					memset(randomText, 0, MAX_PATH);


					outExcelFile.excelDataSet(readText, L"A", read_count+1);

					wsprintf(randomText, L"%c%c%c%c", randomchr(), randomchr(), randomchr(), randomchr());
					outExcelFile.excelDataSet(randomText, L"B", read_count + 1);
				}

				outExcelFile.excelsaveas(1, inReadFilePathTemp);
				outExcelFile.excelclosefile();
				outExcelFile.excelquit();
			}
			
			MessageBox(hDlg, L"작업완료", L"완료!", MB_OK);
			fclose(fp);
			break;
		case IDCANCEL:
			EndDialog(hDlg, TRUE);
			break;
		}
		break;
	case WM_DROPFILES:
		dragFileTimes = 0;
		memset(filenamefromdrag, 0, MAX_PATH);
		hDrop = (HDROP)wParam;
		dragFileTimes = DragQueryFile(hDrop, -1, filenamefromdrag, MAX_PATH);
		for (int fileNameCount = 0; fileNameCount < dragFileTimes; fileNameCount++) {
			DragQueryFile(hDrop, fileNameCount, filenamefromdrag, MAX_PATH);
			if  (wcsstr(filenamefromdrag, L".txt") ){
				SendMessage(IDC_GETFILELIST_H, LB_ADDSTRING, fileNameCount, (LPARAM)filenamefromdrag);
			}
			else
				MessageBox(hDlg, filenamefromdrag, L"This File is Not Text File", 0);
		}
		break;

	case WM_DESTROY:
		processAllKill(hDlg, L"EXCEL.EXE");
		break;
	}
	return FALSE;
}



int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM /*lParam*/, LPARAM lpData) {
	if (uMsg == BFFM_INITIALIZED)
	{
		//BROWSEINFO.lParam에서 설정 해준 값이 lpData로넘어온다.
		// LPARAM으로 path를 넘겨주려면 WParam을 TRUE로,
		// PIDL을 넘겨주려면 FALSE로 넘겨준다.
		if (lpData)
			SendMessage(hwnd, BFFM_SETSELECTION, (WPARAM)TRUE, (LPARAM)lpData);
	}
	return 0;
}


int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPWSTR    lpCmdLine,
	_In_ int       nCmdShow)
{
	DialogBoxW(hInstance, (LPWSTR)IDD_DIALOG1, HWND_DESKTOP, DialogProc);
	return (int)1;
}


BOOL processAllKill(HWND hDlg, const WCHAR* szProcessName)
{
	//MessageBox(hDlg, L"plz exit all excel", L"alert!", MB_OK);
	HANDLE hndl = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	DWORD dwsma = GetLastError();
	HANDLE hHandle;

	DWORD dwExitCode = 0;

	PROCESSENTRY32  procEntry = { 0 };
	procEntry.dwSize = sizeof(PROCESSENTRY32);
	Process32First(hndl, &procEntry);
	while (1)
	{
		if (!wcscmp(procEntry.szExeFile, szProcessName))
		{

			hHandle = ::OpenProcess(PROCESS_ALL_ACCESS, 0, procEntry.th32ProcessID);

			if (::GetExitCodeProcess(hHandle, &dwExitCode))
			{
				if (!::TerminateProcess(hHandle, dwExitCode))
				{
					return FALSE;
				}
			}
		}
		if (!Process32Next(hndl, &procEntry))
		{
			return TRUE;
		}
	}


	return TRUE;
}

unsigned char randomchr() {
	return (rand() % 93) + 33;
}


void OnDrawItem(HWND ah_dlg, UINT a_ctrl_id, DRAWITEMSTRUCT *ap_dis, unsigned int IDC_INT)
{
	if (a_ctrl_id == IDC_INT) {   // '친구 정보 삭제' 버튼을 다시 그린다. 
		// ap_dis->hDC, ap_dis->rcItem을 그대로 사용해도 되지만 저는 설명을 편하게  
		// 하려고 새로 변수를 만들고 대입하겠습니다. 
		HDC h_dc = ap_dis->hDC;  // 버튼 윈도우에 그림을 그리기 위한 DC 핸들 값 
		RECT r = ap_dis->rcItem; // 버튼 영역의 좌표 값  

		// 빨간색 계열의 색상을 사용하여 Brush 객체를 생성한다. 
		HBRUSH h_brush = ::CreateSolidBrush(RGB(192, 64, 32));
		// 버튼 영역 전체를 빨간색 Brush 객체를 사용하여 사각형으로 채운다. 
		::FillRect(h_dc, &r, h_brush);
		// 빨간색 Brush 객체를 제거한다. 
		::DeleteObject(h_brush);

		// 어두운 빨간색 Pen 객체를 생성한다. (바깥쪽 테두리에 사용) 
		HPEN h_out_border_pen = ::CreatePen(PS_SOLID, 1, RGB(64, 0, 0));
		// 밝은 빨간색 Pen 객체를 생성한다. (안쪽 테두리에 사용) 
		HPEN h_in_border_pen = ::CreatePen(PS_SOLID, 1, RGB(255, 112, 82));
		// 사각형 내부를 채우지 않도록 DC에 NULL_BRUSH를 설정한다. 
		HGDIOBJ h_old_brush = ::SelectObject(h_dc, ::GetStockObject(NULL_BRUSH));
		// 바깥쪽 테두리를 만들기 위해서 어두운 빨간색 Pen 객체를 DC 연결한다. 
		HGDIOBJ h_old_pen = ::SelectObject(h_dc, h_out_border_pen);
		// 내부가 채워지지 않는 사각형을 버튼 크기로 그린다. 
		::Rectangle(h_dc, r.left, r.top, r.right, r.bottom);

		if (ap_dis->itemState & ODS_SELECTED) { // 버튼이 눌러진 경우 
			// 버튼이 눌러진 효과를 높이기 위해서 버튼의 Caption을 조금 아래쪽으로 
			// 이동시켜서 출력한다. 
			r.top += 2;
			r.left += 2;
		}
		else {
			// 밝은 빨간색 Pen 객체를 DC에 연결한다. 
			::SelectObject(h_dc, h_in_border_pen);
			// 어두운 빨간색으로 그려진 테두리 안쪽에 다시 테두리를 한 번 더 그린다. 
			::Rectangle(h_dc, r.left + 1, r.top + 1, r.right - 1, r.bottom - 1);
		}
		// 이전에 사용하던 Pen 객체를 DC에 다시 연결한다. 
		::SelectObject(h_dc, h_old_pen);

		// 밝은 빨간색 Pen 객체를 제거한다. 
		::DeleteObject(h_in_border_pen);
		// 어두운 빨간색 Pen 객체를 제거한다. 
		::DeleteObject(h_out_border_pen);

		wchar_t name[400];
		// 버튼 컨트롤에서 Caption 정보를 복사한다. 
		//int length = ::GetDlgItemText(ah_dlg, IDC_INT, name, 400);
		int length = GetWindowText(ah_dlg, name, 400);
		// 배경 그리기 모드를 투명화 모드로 설정한다. 
		int old_mode = ::SetBkMode(h_dc, TRANSPARENT);
		// 글자 색상을 흰색에 가까운 노란색으로 설정한다. 
		::SetTextColor(h_dc, RGB(255, 255, 200));
		// 버튼의 Caption 정보를 버튼 영역의 가운데에 출력한다. 
		::DrawText(h_dc, name, length, &r, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
		// 배경 그리기 모드를 기존 모드로 되돌린다. 
		::SetBkMode(h_dc, old_mode);
	}
}
