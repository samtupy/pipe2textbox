/*
	A purely WINAPI based program which queues any text piped to it for display in a multi-line text box so that it can be easily reviewed with a screen reader.
	This was created as a workaround for the inability to scroll up in a terminal window using NVDA. Instead of running "command --help|clip" and then pasting that result into notepad for review, this purposefully tiny application does that in one step.
	It's worth noting that partially for extra challenge and partially to make the program as utterly small as possible, we do not link with a crt here (This is PURE win32 API)!
	
	Copyright (c) 2023 Sam Tupy, under the MIT license (see the file license).
*/

#define UNICODE

#include <windows.h>
#include <richedit.h>
#include <strsafe.h>
#include "textbox.h"

#define TXTBIT_ALLOWBEEP 0x00000800

typedef struct ITextServicesVtbl {
	HRESULT (STDMETHODCALLTYPE* QueryInterface)(void* This, REFIID riid, void** ppvObject);
	ULONG (STDMETHODCALLTYPE* AddRef)(void* This);
	ULONG (STDMETHODCALLTYPE* Release)(void* This);
	HRESULT (STDMETHODCALLTYPE* TxSendMessage)(void* This, UINT msg, WPARAM wparam, LPARAM lparam, LRESULT* plresult);
	HRESULT (STDMETHODCALLTYPE* TxDraw)(void* This, DWORD dwDrawAspect, LONG lindex, void* pvAspect, DVTARGETDEVICE* ptd, HDC hdcDraw, HDC hicTargetDev, LPCRECTL lprcBounds, LPCRECTL lprcWBounds, LPRECT lprcUpdate, BOOL (CALLBACK* pfnContinue)(DWORD), DWORD dwContinue, LONG lViewId);
	HRESULT (STDMETHODCALLTYPE* TxGetHScroll)(void* This, LONG* plMin, LONG* plMax, LONG* plPos, LONG* plPage, BOOL* pfEnabled);
	HRESULT (STDMETHODCALLTYPE* TxGetVScroll)(void* This, LONG* plMin, LONG* plMax, LONG* plPos, LONG* plPage, BOOL* pfEnabled);
	HRESULT (STDMETHODCALLTYPE* OnTxSetCursor)(void* This, DWORD dwDrawAspect, LONG lindex, void* pvAspect, DVTARGETDEVICE* ptd, HDC hdcDraw, HDC hicTargetDev, LPCRECT lprcClient, INT x, INT y);
	HRESULT (STDMETHODCALLTYPE* TxQueryHitPoint)(void* This, DWORD dwDrawAspect, LONG lindex, void* pvAspect, DVTARGETDEVICE* ptd, HDC hdcDraw, HDC hicTargetDev, LPCRECT lprcClient, INT x, INT y, DWORD* pHitResult);
	HRESULT (STDMETHODCALLTYPE* OnTxInPlaceActivate)(void* This, LPCRECT prcClient);
	HRESULT (STDMETHODCALLTYPE* OnTxInPlaceDeactivate)(void* This);
	HRESULT (STDMETHODCALLTYPE* OnTxUIActivate)(void* This);
	HRESULT (STDMETHODCALLTYPE* OnTxUIDeactivate)(void* This);
	HRESULT (STDMETHODCALLTYPE* TxGetText)(void* This, BSTR* pbstrText);
	HRESULT (STDMETHODCALLTYPE* TxSetText)(void* This, LPCWSTR pszText);
	HRESULT (STDMETHODCALLTYPE* TxGetCurTargetX)(void* This, LONG* px);
	HRESULT (STDMETHODCALLTYPE* TxGetBaseLinePos)(void* This, LONG* py);
	HRESULT (STDMETHODCALLTYPE* TxGetNaturalSize)(void* This, DWORD dwAspect, HDC hdcDraw, HDC hicTargetDev, DVTARGETDEVICE* ptd, DWORD dwMode, const SIZEL* psizelExtent, LONG* pwidth, LONG* pheight);
	HRESULT (STDMETHODCALLTYPE* TxGetDropTarget)(void* This, IDropTarget** ppDropTarget);
	HRESULT (STDMETHODCALLTYPE* OnTxPropertyBitsChange)(void* This, DWORD dwMask, DWORD dwBits);
	HRESULT (STDMETHODCALLTYPE* TxGetCachedSize)(void* This, DWORD* pdwWidth, DWORD* pdwHeight);
} ITextServicesVtbl;

typedef struct ITextServices {
	ITextServicesVtbl* lpVtbl;
} ITextServices;

// forward declarations
BOOL CALLBACK textbox_callback(HWND hwnd, UINT message, WPARAM wp, LPARAM lp);
LRESULT CALLBACK edit_control_callback(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
void disable_richedit_beeps(HMODULE richedit_module, HWND richedit_control);
void find(HWND hwnd, int dir);
void save(HWND hwnd);
// Globals required for the find dialog.
WNDPROC original_edit_control_callback = NULL; // We need to subclass the textbox since the main dialog isn't receiving WM_KEYDOWN for some reason.
wchar_t text_to_search[256];
HWND find_dlg = NULL;
DWORD find_dlg_flags = 0; // If not global, f3 and shift+f3 won't work correctly regarding case sensativity and whole word searches.
UINT M_FINDMSGSTRING;

int main() {
	HANDLE cin, cout, process_heap;
	HWND dlg, output_box;
	HMODULE richedit_module;
	MSG msg;
	size_t cursor, allocated, text_transcode_result;
	DWORD console_type, console_bytes_read, text_codepage;
	wchar_t* output, * output_adjusted;
	char output_tmp[2052]; // Need 4 extra bytes for proper reading of UTF8 data.
	cin = GetStdHandle(STD_INPUT_HANDLE);
	if((console_type = GetFileType(cin)) == FILE_TYPE_CHAR) { // We don't want to wait for user input when reading stdin if there is no pipe.
		const char* message = "This tool displays any text piped to it in a multi-line input box for screen reader accessibility. You should either run \"command | show\" or \"show < filename\".\r\n";
		cout = GetStdHandle(STD_OUTPUT_HANDLE);
		if(cout != INVALID_HANDLE_VALUE)
			WriteFile(cout, message, lstrlenA(message), NULL, NULL);
		return 0;
	} else if(console_type == FILE_TYPE_PIPE) text_codepage = CP_UTF8;
	else text_codepage = 0;
	// Allocate a buffer to store characters from stdin
	allocated = 4096; cursor = 0;
	process_heap = GetProcessHeap();
	output = (wchar_t*)HeapAlloc(process_heap, 0, allocated);
	// Read until we reach the end of the stdin stream.
	while(ReadFile(cin, output_tmp, 2048, &console_bytes_read, NULL) && console_bytes_read) {
		// Reallocate the output buffer if needed.
		if((cursor + console_bytes_read) * sizeof(wchar_t) >= allocated) {
			allocated *= 2;
			output = (wchar_t*)HeapReAlloc(process_heap, 0, output, allocated);
		}
		if(!text_codepage) { // certainly a file then
			text_codepage = IsTextUnicode(output_tmp, console_bytes_read, NULL);
			if(!text_codepage) text_codepage = CP_UTF8;
		}
		if(text_codepage == 1) { // Probably unicode (only check that in first text block)
			StringCchCopyW(output + cursor, (console_bytes_read / 2) + 1, (wchar_t*)&output_tmp);
			cursor += console_bytes_read / 2;
			continue;
		}
		// If we're reading UTF8 data, we can't just blindly transcode after reading 2048 bytes exactly. Some UTF8 magic numbers here and some code I'm not very proud of (spent hours getting it right erm I mean finding just a few small mistakes), sorry. Probably I should try making some part of this it's own function later, too tired now. Even if text isn't actually UTF8, only a couple of extra bytes should be read in the worst case.
		if(text_codepage == CP_UTF8 &&console_bytes_read < 2052 && (output_tmp[console_bytes_read -1] & (1 << 7)) != 0) {
			DWORD character_size, bytes_read_tmp, i;
			character_size = 0;
			for(i = 1; i <= 4; i ++) { // One of these 4 bytes will tell us the size of the character we're dealing with.
				unsigned char c;
				c = output_tmp[console_bytes_read - i];
				if(c < 192) continue; // another continuation char.
				for(character_size = 2; character_size <=4; character_size ++) {
					if((c & (1 << (7 - character_size))) == 0) break;
				}
				if(character_size) break;
			}
			if(character_size - i > 0) {
				ReadFile(cin, output_tmp + console_bytes_read, character_size - i, &bytes_read_tmp, NULL);
				console_bytes_read += bytes_read_tmp;
			}
		} // fwiw, that UTF8 correction thoroughly sucked.
		text_transcode_result = MultiByteToWideChar(text_codepage != -1? text_codepage : 1252, text_codepage == CP_UTF8 ? 0 : (text_codepage != -1? MB_ERR_INVALID_CHARS : 0), output_tmp, console_bytes_read, output + cursor, console_bytes_read); // Codepage can end up being unquestioningly windows1252 if transcoding attempts below fail, but we don't want to re-check encodings every text block.
		if(!text_transcode_result) {
			text_codepage = CP_ACP;
			text_transcode_result = MultiByteToWideChar(text_codepage, 0, output_tmp, console_bytes_read, output + cursor, console_bytes_read);
		} if(!text_transcode_result) { // We are truly desperate...
			text_codepage = -1; // Gotta choose something... in this case windows-1252 as seen above and below.
			text_transcode_result = MultiByteToWideChar(1252, 0, output_tmp, console_bytes_read, output + cursor, console_bytes_read);
		} if(!text_transcode_result) { // and now we are truly done for.
			MessageBox(0, L"Failed to decode this input", L"Error", MB_ICONERROR);
			HeapFree(process_heap, 0, output);
			ExitProcess(1);
		}
		cursor += text_transcode_result;
	}
	if(!cursor) {
		HeapFree(process_heap, 0, output);
		return 0; // No STDOutput, best to bail out encase stderr has anything to say.
	}
	output[cursor] = '\0'; // Making sure the output string is NULL terminated.
	output_adjusted = output; // As you can see below we may need to skip characters at the beginning of text for some reason, must retain the output pointer though so it can be freed.
	// Check for and remove a byte order mark in the case of a UTF16 file being piped.
	if(output[0] == 0xfeff) output_adjusted += 1;
	// Prepare and show the dialog.
	richedit_module = LoadLibrary(L"MSFTEDIT.dll"); // This will register the MSFTEDIT_CLASS.
	if(!richedit_module) {
		HeapFree(process_heap, 0, output);
		return 1;
	}
	dlg = CreateDialog(NULL, MAKEINTRESOURCE(textbox), 0, (DLGPROC)textbox_callback);
	if(!dlg) {
		HeapFree(process_heap, 0, output);
		return 1;
	}
	output_box = GetDlgItem(dlg, IDC_TEXT);
	disable_richedit_beeps(richedit_module, output_box);
	SendMessage(output_box, EM_SETLIMITTEXT, 0, 0);
	SetWindowText(output_box, output_adjusted);
	SendMessage(output_box, EM_SETSEL, 0, 0);
	original_edit_control_callback = (WNDPROC)SetWindowLongPtr(output_box, GWLP_WNDPROC, (LONG_PTR)edit_control_callback);
	HeapFree(process_heap, 0, output);
	// A couple tiny things for the find text dialog.
	RtlSecureZeroMemory(&text_to_search, sizeof(text_to_search));
	M_FINDMSGSTRING = RegisterWindowMessage(FINDMSGSTRING);
	// Finally, handle events.
	while(GetMessage(&msg, 0, 0, 0)) {
		if(find_dlg && IsDialogMessage(find_dlg, &msg)) continue;
		else if(IsDialogMessage(dlg, &msg)) continue;
		else {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	ExitProcess(msg.wParam); // We're using /nodefaultlib in the linker as we don't need the crt for this application, forcing us to call ExitProcess manually to prevent an app hang.
	return msg.wParam; // LOL if this line of code executes then ExitProcess somehow failed, fancy that.
}

BOOL CALLBACK textbox_callback(HWND hwnd, UINT message, WPARAM wp, LPARAM lp) {
	int control, event;
	switch(message) {
	case WM_COMMAND: {
		control = LOWORD(wp);
		event = HIWORD(wp);
		if(control == IDCANCEL && event == BN_CLICKED)
			DestroyWindow(hwnd);
		else if(control == IDC_FIND && event == BN_CLICKED)
			find(GetDlgItem(hwnd, IDC_TEXT), 0);
		else if(control == IDC_SAVE && event == BN_CLICKED)
			save(GetDlgItem(hwnd, IDC_TEXT));
		return TRUE;
	}
	case WM_DESTROY: {
		PostQuitMessage(0);
		return TRUE;
	}
	}
	return FALSE;
}

// This function makes richedit controls stop making a sound if you try to scroll past their borders. Thanks to https://stackoverflow.com/questions/55884687/how-to-eliminate-the-messagebeep-from-the-richedit-control for the original C++ code.
void disable_richedit_beeps(HMODULE richedit_module, HWND richedit_control) {
	IUnknown* unknown = NULL;
	ITextServices* ts = NULL;
	IID* ITextservicesId = (IID*)GetProcAddress(richedit_module, "IID_ITextServices");
	if (!ITextservicesId) return;
	if (!SendMessage(richedit_control, EM_GETOLEINTERFACE, 0, (LPARAM)&unknown) || !unknown) return;
	HRESULT hr = unknown->lpVtbl->QueryInterface(unknown, ITextservicesId, (void**)&ts);
	unknown->lpVtbl->Release(unknown);
	if (FAILED(hr) || !ts) return;
	ts->lpVtbl->OnTxPropertyBitsChange(ts, TXTBIT_ALLOWBEEP, 0);
	ts->lpVtbl->Release(ts);
}

// This implements the find dialog.
void find(HWND hwnd, int dir) {
	if(dir == 0 || !text_to_search[0]) {
		static FINDREPLACE fr; // If this is not global or static the program will crash as soon as the find dialog tries sending it's first message.
		RtlSecureZeroMemory(&fr, sizeof(fr));
		fr.lStructSize = sizeof(fr);
		fr.hwndOwner = hwnd;
		fr.lpstrFindWhat = text_to_search;
		fr.wFindWhatLen = 256;
		fr.Flags = (dir >= 0? FR_DOWN : 0) | (find_dlg_flags & FR_MATCHCASE? FR_MATCHCASE : 0) | (find_dlg_flags & FR_WHOLEWORD ? FR_WHOLEWORD : 0); // Maybe this is clunky? I'm on the fense as to whether fr.Flags = find_dlg_flags then manually resetting the FR_DOWN flag as needed would be less so.
		find_dlg = FindText(&fr);
	} else {
		FINDTEXTEXW ft;
		RtlSecureZeroMemory(&ft, sizeof(ft));
		SendMessage(hwnd, EM_EXGETSEL, 0, (LPARAM)&ft.chrg);
		if(ft.chrg.cpMin > 0) ft.chrg.cpMin += dir; // We want to start searching at the next or previous cursor position, not the current one.
		ft.chrg.cpMax = -1;
		ft.lpstrText = text_to_search;
		SendMessage(hwnd, EM_FINDTEXTEX, (dir > 0? FR_DOWN : 0) | (find_dlg_flags & FR_MATCHCASE? FR_MATCHCASE : 0) | (find_dlg_flags & FR_WHOLEWORD ? FR_WHOLEWORD : 0), (LPARAM)&ft);
		if(ft.chrgText.cpMin >= 0) {
			SendMessage(hwnd, EM_EXSETSEL, 0, (LPARAM)&ft.chrgText);
			SetFocus(hwnd); // Encase find dialog was activated through button instead of shortcut.
			SendMessage(find_dlg, WM_CLOSE, 0, 0);
		} else MessageBox((find_dlg? find_dlg : hwnd), L"nothing found for the given search", L"error", MB_ICONERROR);
	}
}

// we don't want the user to have to copy the text they're seeing to notepad just because they want to preserve it, that defeats some of the point of this thing!
// Callback that writes data from a text field to a file.
DWORD CALLBACK save_editstream_callback(DWORD_PTR cookie,LPBYTE buffer, LONG bufsize, LONG* bytes_written) {
	DWORD dw_bytes_written;
	HANDLE* file=(HANDLE*)cookie;
	if(WriteFile(*file, buffer, bufsize, &dw_bytes_written, NULL)) {
	*bytes_written = dw_bytes_written;
		return 0;
	}
	else return 1;
}
void save(HWND hwnd) {
	EDITSTREAM editstream;
	wchar_t save_path[MAX_PATH];
	HANDLE save_file;
	OPENFILENAME ofn;
	StringCchCopyW(save_path, MAX_PATH, L"output.txt");
	RtlSecureZeroMemory(&ofn, sizeof(ofn));
	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = hwnd;
	ofn.lpstrFile = save_path;
	ofn.nMaxFile = MAX_PATH;
	ofn.lpstrFilter = L"TXT files (*.txt)\0*.txt\0All files\0*.*\0\0";
	ofn.nFilterIndex = 1;
	ofn.lpstrInitialDir = L"";
	ofn.Flags = OFN_OVERWRITEPROMPT;
	if(!GetSaveFileName(&ofn)) return;
	save_file = CreateFile(save_path, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if(!save_file) return;
	RtlSecureZeroMemory(&editstream, sizeof(editstream));
	editstream.dwCookie = (DWORD_PTR)&save_file;
	editstream.pfnCallback = save_editstream_callback;
	SendMessage(hwnd, EM_STREAMOUT, (CP_UTF8 << 16) | SF_USECODEPAGE | SF_TEXT, (LPARAM)&editstream);
	CloseHandle(save_file);
	if(editstream.dwError) MessageBox(hwnd, L"potential error saving", L"warning", 0);
}

LRESULT CALLBACK edit_control_callback(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	FINDREPLACE* fr;
	BOOL ret = TRUE; // No switch case here because M_FINDMSGSTRING is not a constant.
	if(msg == WM_KEYDOWN) {
	if(wParam == 'F' && GetAsyncKeyState(VK_CONTROL) & 0x8000)
		find(hwnd, 0);
	else if(wParam == 'S' && GetAsyncKeyState(VK_CONTROL) & 0x8000)
		save(hwnd);
	else if(wParam == VK_F3)
		find(hwnd, GetAsyncKeyState(VK_SHIFT) & 0x8000? -1 : 1);
	else ret = FALSE;
	if(ret) return ret;
	} else if(msg == M_FINDMSGSTRING) {
		fr = (FINDREPLACE*)lParam;
		if(fr->Flags & FR_DIALOGTERM) {
			find_dlg = NULL;
			return TRUE;
		}
		find_dlg_flags = fr->Flags;
		if(fr->Flags & FR_FINDNEXT) find(hwnd, fr->Flags & FR_DOWN? 1 : -1);
		return TRUE;
	}
	return CallWindowProc(original_edit_control_callback, hwnd, msg, wParam, lParam);
}
