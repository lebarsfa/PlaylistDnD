#pragma once

#include "resource.h"
#include "Utils.h"
#include "M3UUtils.h"
#include <random>
#include <shlobj.h>      // FILEGROUPDESCRIPTOR, FILEDESCRIPTOR
#include <shlwapi.h>     // SHCreateStreamOnFileEx
#include <objidl.h>      // IStream
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")

// Globals
HINSTANCE g_hInst = NULL;
HICON g_hIconBig = NULL;
HICON g_hIconSmall = NULL;
HFONT hLabelFont;
HWND g_hLabel = NULL;
HWND g_hList = NULL;
HWND g_hHeader = NULL;
IDropTarget* g_dropTarget = nullptr;
int g_sortColumn = 0;
bool g_sortAscending = true;

struct DragTrackState {
	bool tracking = false;
	POINT downPt = { 0,0 }; // client coords
};

DragTrackState s_dt;

// button id
#define ID_BTN_SHUFFLE 4001

// Shuffle button handle
HWND g_hShuffleBtn = NULL;

// Dark theme colors and brushes
COLORREF g_clrBg = RGB(28, 28, 36);// RGB(18, 24, 48);        // night blue / violet
COLORREF g_clrCtrl = RGB(28, 28, 36);// RGB(36, 36, 48);// RGB(18, 24, 48);// RGB(28, 34, 64);      // control background (slightly lighter)
COLORREF g_clrText = RGB(240, 235, 220);// RGB(230, 230, 250);   // light text
COLORREF g_clrEvenCol = RGB(44, 44, 64);// RGB(36, 44, 84);
COLORREF g_clrOddCol = RGB(48, 60, 80);// RGB(42, 64, 104);
COLORREF g_clrSelBg = RGB(70, 110, 140);// RGB(60, 120, 180);
COLORREF g_clrSelText = RGB(255, 255, 255);// RGB(255, 255, 255);
COLORREF gridColor = RGB(0, 0, 0);// RGB(48, 56, 96);
COLORREF gridColorVertical = RGB(0, 0, 0);// RGB(44, 48, 72);// RGB(40, 48, 88);
COLORREF g_clrScrollBg = RGB(36, 40, 56);// RGB(36, 44, 64);    // scrollbar track
COLORREF g_clrScrollThumb = RGB(90, 110, 130);// RGB(80, 120, 180); // thumb
HBRUSH g_hbrBg = NULL;
HBRUSH g_hbrCtrl = NULL;
HBRUSH g_hbrScrollBg = NULL;
HBRUSH g_hbrScrollThumb = NULL;

// Track which columns were resized by the user
bool g_colUserSized[] = { false, false, false, false };

// Enumerate files in a directory (non-recursive). If filterExts is empty, add all files.
inline void EnumerateDirectoryFiles(const std::wstring& dir, std::vector<std::wstring>& out,
	const std::vector<std::wstring>& filterExts = {}) {
	std::wstring search = dir;
	if (!search.empty() && search.back() != L'\\') search.push_back(L'\\');
	search += L"*";

	WIN32_FIND_DATAW fd;
	HANDLE h = FindFirstFileW(search.c_str(), &fd);
	if (h == INVALID_HANDLE_VALUE) return;
	do {
		if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
		if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue; // skip subdirs (non-recursive)
		std::wstring full = dir;
		if (!full.empty() && full.back() != L'\\') full.push_back(L'\\');
		full += fd.cFileName;
		if (filterExts.empty()) {
			out.push_back(full);
		}
		else {
			size_t pos = full.find_last_of(L'.');
			std::wstring ext = (pos == std::wstring::npos) ? L"" : full.substr(pos);
			std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
			for (auto& fe : filterExts) {
				if (ext == fe) { out.push_back(full); break; }
			}
		}
	} while (FindNextFileW(h, &fd));
	FindClose(h);
}

// create a DPI-aware font (call once, e.g., in WM_CREATE)
inline HFONT CreateUiFont(HWND hwnd, const wchar_t* faceName, int pointSize)
{
	// get DPI for the window (Windows 10+). Fallback to 96 if unavailable.
	UINT dpi = 96;
	HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
	typedef UINT(WINAPI* GetDpiForWindow_t)(HWND);
	GetDpiForWindow_t pGetDpiForWindow = (GetDpiForWindow_t)GetProcAddress(hUser32, "GetDpiForWindow");
	if (pGetDpiForWindow) dpi = pGetDpiForWindow(hwnd);
	else {
		HDC hdc = GetDC(hwnd);
		if (hdc) { dpi = GetDeviceCaps(hdc, LOGPIXELSY); ReleaseDC(hwnd, hdc); }
	}

	LOGFONTW lf = {};
	lf.lfHeight = -MulDiv(pointSize, (int)dpi, 72); // negative for character height
	lf.lfWeight = FW_NORMAL;
	lf.lfCharSet = DEFAULT_CHARSET;
	lf.lfQuality = CLEARTYPE_QUALITY;
	wcsncpy_s(lf.lfFaceName, faceName, ARRAYSIZE(lf.lfFaceName)-1);
	return CreateFontIndirectW(&lf);
}

inline std::wstring FormatIdToName(UINT cf)
{
	wchar_t name[256] = {};
	if (cf < 0xC000) {
		switch (cf) {
		case CF_HDROP: return L"CF_HDROP";
		case CF_TEXT: return L"CF_TEXT";
		case CF_UNICODETEXT: return L"CF_UNICODETEXT";
		default: break;
		}
	}
	if (GetClipboardFormatNameW(cf, name, ARRAYSIZE(name)) > 0) {
		return std::wstring(name);
	}
	wchar_t buf[64];
	swprintf_s(buf, L"CF_0x%04X", cf);
	return std::wstring(buf);
}

// Replace your CreateShellIDListHGlobal with this implementation
inline HGLOBAL CreateShellIDListHGlobal(const std::vector<std::wstring>& paths)
{
	if (paths.empty()) return NULL;

	// compute common parent folder
	std::wstring parent = paths[0];
	size_t pos = parent.find_last_of(L"\\/");
	if (pos == std::wstring::npos) parent = L"";
	else parent = parent.substr(0, pos);

	// get PIDL for parent folder
	LPITEMIDLIST pidlParent = nullptr;
	HRESULT hr = SHParseDisplayName(parent.c_str(), NULL, &pidlParent, 0, NULL);
	if (FAILED(hr) || !pidlParent) {
		// fallback: try desktop as parent (empty)
		hr = SHGetDesktopFolder(nullptr);
		// if we can't get parent, still try to build CIDA with item-only PIDLs (less ideal)
	}

	// collect item SHITEMIDs (last ID only) and their sizes
	struct ItemBuf { BYTE* data; SIZE_T size; };
	std::vector<ItemBuf> items;
	items.reserve(paths.size());

	for (auto& p : paths) {
		LPITEMIDLIST pidlFull = nullptr;
		if (SUCCEEDED(SHParseDisplayName(p.c_str(), NULL, &pidlFull, 0, NULL)) && pidlFull) {
			// find last SHITEMID within pidlFull
			PBYTE pb = (PBYTE)pidlFull;
			// walk to last SHITEMID
			PBYTE last = pb;
			while (true) {
				// each SHITEMID: USHORT cb; followed by cb bytes; terminator is 0 USHORT
				USHORT cb = *(USHORT*)last;
				if (cb == 0) break;
				PBYTE next = last + cb;
				// check if next is terminator
				USHORT nextcb = *(USHORT*)next;
				if (nextcb == 0) {
					// last currently points to the last SHITEMID
					break;
				}
				last = next;
			}
			// size of that single SHITEMID (cb + sizeof(cb) already included)
			USHORT lastcb = *(USHORT*)last;
			SIZE_T itemSize = lastcb + sizeof(USHORT); // include terminating size field
			// allocate buffer and copy the single SHITEMID plus terminating 0 WORD
			BYTE* buf = (BYTE*)CoTaskMemAlloc(itemSize + sizeof(USHORT)); // ensure space for terminator
			if (buf) {
				memcpy(buf, last, itemSize);
				// append a terminating zero WORD if not present
				USHORT term = 0;
				memcpy(buf + itemSize, &term, sizeof(term));
				items.push_back({ buf, itemSize + sizeof(term) });
			}
			CoTaskMemFree(pidlFull);
		}
		else {
			// if parsing failed, push an empty item (single zero WORD)
			BYTE* buf = (BYTE*)CoTaskMemAlloc(sizeof(USHORT));
			if (buf) {
				USHORT term = 0;
				memcpy(buf, &term, sizeof(term));
				items.push_back({ buf, sizeof(USHORT) });
			}
		}
	}

	// parent PIDL bytes size
	SIZE_T parentSize = 0;
	BYTE* parentBytes = nullptr;
	if (pidlParent) {
		parentSize = ILGetSize(pidlParent);
		parentBytes = (BYTE*)pidlParent; // will free later with CoTaskMemFree
	}
	else {
		// empty parent -> single zero WORD
		parentSize = sizeof(USHORT);
		parentBytes = (BYTE*)CoTaskMemAlloc(parentSize);
		if (parentBytes) {
			USHORT term = 0;
			memcpy(parentBytes, &term, sizeof(term));
		}
	}

	UINT cItems = (UINT)items.size();
	SIZE_T headerSize = sizeof(UINT) + (cItems + 1) * sizeof(UINT); // cidl + offsets[0..cItems]
	SIZE_T totalPidlsSize = parentSize;
	for (auto& it : items) totalPidlsSize += it.size;
	SIZE_T total = headerSize + totalPidlsSize;

	HGLOBAL hg = GlobalAlloc(GHND | GMEM_SHARE, total);
	if (!hg) {
		if (pidlParent) CoTaskMemFree(pidlParent);
		else if (parentBytes) CoTaskMemFree(parentBytes);
		for (auto& it : items) if (it.data) CoTaskMemFree(it.data);
		return NULL;
	}

	BYTE* mem = (BYTE*)GlobalLock(hg);
	if (!mem) {
		GlobalFree(hg);
		if (pidlParent) CoTaskMemFree(pidlParent);
		else if (parentBytes) CoTaskMemFree(parentBytes);
		for (auto& it : items) if (it.data) CoTaskMemFree(it.data);
		return NULL;
	}

	// write header
	UINT* pu = (UINT*)mem;
	pu[0] = cItems;
	// offsets: aoffset[0] = offset to parent PIDL (immediately after header)
	SIZE_T cur = headerSize;
	pu[1] = (UINT)cur;
	// copy parent PIDL bytes
	memcpy(mem + cur, parentBytes, parentSize);
	cur += parentSize;

	// offsets for each item
	for (UINT i = 0; i < cItems; ++i) {
		pu[2 + i] = (UINT)cur;
		memcpy(mem + cur, items[i].data, items[i].size);
		cur += items[i].size;
	}

	GlobalUnlock(hg);

	// free temporary buffers
	if (pidlParent) CoTaskMemFree(pidlParent);
	else if (parentBytes) CoTaskMemFree(parentBytes);
	for (auto& it : items) if (it.data) CoTaskMemFree(it.data);

	return hg;
}

class FileDropDataObject : public IDataObject {
	LONG _ref;
	HGLOBAL _hDrop;
	std::vector<std::wstring> _paths;

public:
	FileDropDataObject(const std::vector<std::wstring>& paths) : _ref(1), _hDrop(NULL), _paths(paths) {
		// build CF_HDROP HGLOBAL (Unicode)
		size_t totalWchars = 0;
		for (auto& p : paths) totalWchars += p.size() + 1;
		totalWchars += 1;
		SIZE_T bytes = sizeof(DROPFILES) + totalWchars * sizeof(wchar_t);
		_hDrop = GlobalAlloc(GHND | GMEM_SHARE, bytes);
		if (!_hDrop) return;

		BYTE* mem = (BYTE*)GlobalLock(_hDrop);
		if (!mem) { GlobalFree(_hDrop); _hDrop = NULL; return; }

		DROPFILES* df = (DROPFILES*)mem;
		ZeroMemory(df, sizeof(DROPFILES));
		df->pFiles = sizeof(DROPFILES);
		df->fWide = TRUE;

		wchar_t* wptr = (wchar_t*)(mem + sizeof(DROPFILES));
		wchar_t* cur = wptr;
		for (auto& p : paths) {
			memcpy(cur, p.c_str(), (p.size() + 1) * sizeof(wchar_t));
			cur += p.size() + 1;
		}
		*cur = L'\0';
		GlobalUnlock(_hDrop);
	}

	~FileDropDataObject() {
		if (_hDrop) GlobalFree(_hDrop);
	}

	// IUnknown
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override {
		if (!ppvObject) return E_POINTER;
		if (riid == IID_IUnknown || riid == IID_IDataObject) {
			*ppvObject = static_cast<IDataObject*>(this);
			AddRef();
			return S_OK;
		}
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}
	ULONG STDMETHODCALLTYPE AddRef(void) override { return InterlockedIncrement(&_ref); }
	ULONG STDMETHODCALLTYPE Release(void) override {
		ULONG c = InterlockedDecrement(&_ref);
		if (c == 0) delete this;
		return c;
	}

	// IDataObject

	HRESULT STDMETHODCALLTYPE GetData(FORMATETC* pFormatEtc, STGMEDIUM* pMedium) override {
		if (!pFormatEtc || !pMedium) return E_POINTER;

		// debug log
		wchar_t dbg[512];
		swprintf_s(dbg, L"GetData: cf=%s (0x%08X) tymed=0x%X lindex=%d\n",
			FormatIdToName(pFormatEtc->cfFormat).c_str(),
			pFormatEtc->cfFormat, pFormatEtc->tymed, pFormatEtc->lindex);
		OutputDebugStringW(dbg);

		// CF_HDROP (TYMED_HGLOBAL)
		if ((pFormatEtc->cfFormat == CF_HDROP) && (pFormatEtc->tymed & TYMED_HGLOBAL)) {
			if (!_hDrop) return E_FAIL;
			SIZE_T size = GlobalSize(_hDrop);
			HGLOBAL hg = GlobalAlloc(GHND | GMEM_SHARE, size);
			if (!hg) return E_OUTOFMEMORY;
			void* src = GlobalLock(_hDrop);
			void* dst = GlobalLock(hg);
			if (!src || !dst) {
				if (src) GlobalUnlock(_hDrop);
				if (dst) GlobalUnlock(hg);
				GlobalFree(hg);
				return E_FAIL;
			}
			memcpy(dst, src, size);
			GlobalUnlock(_hDrop);
			GlobalUnlock(hg);
			pMedium->tymed = TYMED_HGLOBAL;
			pMedium->hGlobal = hg;
			pMedium->pUnkForRelease = nullptr;
			return S_OK;
		}

		// CFSTR_INETURLW (UniformResourceLocatorW) - UTF-16 newline-separated file:/// URLs
		UINT cfInetUrlW = RegisterClipboardFormatW(CFSTR_INETURLW);
		if (pFormatEtc->cfFormat == cfInetUrlW && (pFormatEtc->tymed & TYMED_HGLOBAL)) {
			std::wstring urls;
			for (size_t i = 0; i < _paths.size(); ++i) {
				const std::wstring& path = _paths[i];
				std::wstring url = L"file:///";
				for (wchar_t ch : path) url.push_back(ch == L'\\' ? L'/' : ch);
				urls += url;
				if (i + 1 < _paths.size()) urls += L"\r\n";
			}
			SIZE_T bytes = (urls.size() + 1) * sizeof(wchar_t);
			HGLOBAL hg = GlobalAlloc(GHND | GMEM_SHARE, bytes);
			if (!hg) return E_OUTOFMEMORY;
			void* dst = GlobalLock(hg);
			memcpy(dst, urls.c_str(), bytes);
			GlobalUnlock(hg);
			pMedium->tymed = TYMED_HGLOBAL;
			pMedium->hGlobal = hg;
			pMedium->pUnkForRelease = nullptr;
			return S_OK;
		}

		// CF_PREFERREDDROPEFFECT - DWORD with DROPEFFECT_COPY
		UINT cfPref = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
		if (pFormatEtc->cfFormat == cfPref && (pFormatEtc->tymed & TYMED_HGLOBAL)) {
			HGLOBAL hg = GlobalAlloc(GHND | GMEM_SHARE, sizeof(DWORD));
			if (!hg) return E_OUTOFMEMORY;
			void* dst = GlobalLock(hg);
			DWORD effect = DROPEFFECT_COPY;
			memcpy(dst, &effect, sizeof(effect));
			GlobalUnlock(hg);
			pMedium->tymed = TYMED_HGLOBAL;
			pMedium->hGlobal = hg;
			pMedium->pUnkForRelease = nullptr;
			return S_OK;
		}

		// CFSTR_FILEDESCRIPTORW - FILEGROUPDESCRIPTORW (TYMED_HGLOBAL)
		UINT cfFileDesc = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);
		if (pFormatEtc->cfFormat == cfFileDesc && (pFormatEtc->tymed & TYMED_HGLOBAL)) {
			HGLOBAL hg = CreateFileGroupDescriptorHGlobal();
			if (!hg) return E_OUTOFMEMORY;
			pMedium->tymed = TYMED_HGLOBAL;
			pMedium->hGlobal = hg;
			pMedium->pUnkForRelease = nullptr;
			return S_OK;
		}

		// CFSTR_FILECONTENTS - prefer TYMED_ISTREAM (caller may request ISTREAM)
		UINT cfFileCont = RegisterClipboardFormatW(CFSTR_FILECONTENTS);
		if (pFormatEtc->cfFormat == cfFileCont) {
			// debug log the request
			wchar_t dbg2[256];
			swprintf_s(dbg2, L"GetData request: CFSTR_FILECONTENTS tymed=0x%X lindex=%d\n", pFormatEtc->tymed, pFormatEtc->lindex);
			OutputDebugStringW(dbg2);

			LONG idx = pFormatEtc->lindex;
			if (idx < 0 || (size_t)idx >= _paths.size()) return DV_E_LINDEX;

			// If caller asked for ISTREAM, provide it
			if (pFormatEtc->tymed & TYMED_ISTREAM) {
				IStream* pStream = nullptr;
				HRESULT hr = SHCreateStreamOnFileEx(_paths[idx].c_str(),
					STGM_READ | STGM_SHARE_DENY_NONE,
					FILE_ATTRIBUTE_NORMAL, FALSE, NULL, &pStream);
				if (SUCCEEDED(hr) && pStream) {
					pMedium->tymed = TYMED_ISTREAM;
					pMedium->pstm = pStream; // caller will Release()
					pMedium->pUnkForRelease = nullptr;
					return S_OK;
				}
				// if stream creation failed, fall through to HGLOBAL fallback below (if implemented)
			}

			// Optional fallback: if caller accepts HGLOBAL, return file bytes in HGLOBAL (only for small files)
			if (pFormatEtc->tymed & TYMED_HGLOBAL) {
				HANDLE hFile = CreateFileW(_paths[idx].c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
				if (hFile == INVALID_HANDLE_VALUE) return STG_E_FILENOTFOUND;
				LARGE_INTEGER li = {};
				if (!GetFileSizeEx(hFile, &li)) { CloseHandle(hFile); return E_FAIL; }
				SIZE_T bytes = (SIZE_T)li.QuadPart;
				if (bytes == 0) {
					CloseHandle(hFile); // empty file -> return empty HGLOBAL
					HGLOBAL hgEmpty = GlobalAlloc(GHND | GMEM_SHARE, 1);
					if (!hgEmpty) return E_OUTOFMEMORY;
					void* dstEmpty = GlobalLock(hgEmpty);
					*((char*)dstEmpty) = 0;
					GlobalUnlock(hgEmpty);
					pMedium->tymed = TYMED_HGLOBAL;
					pMedium->hGlobal = hgEmpty;
					pMedium->pUnkForRelease = nullptr;
					return S_OK;
				}
				// guard: avoid huge allocations here; prefer ISTREAM for large files
				if (bytes > (1 << 26)) { CloseHandle(hFile); return STG_E_MEDIUMFULL; } // ~64MB guard
				HGLOBAL hg = GlobalAlloc(GHND | GMEM_SHARE, bytes);
				if (!hg) { CloseHandle(hFile); return E_OUTOFMEMORY; }
				void* dst = GlobalLock(hg);
				DWORD read = 0;
				BOOL ok = ReadFile(hFile, dst, (DWORD)bytes, &read, NULL);
				GlobalUnlock(hg);
				CloseHandle(hFile);
				if (!ok || read != bytes) { GlobalFree(hg); return E_FAIL; }
				pMedium->tymed = TYMED_HGLOBAL;
				pMedium->hGlobal = hg;
				pMedium->pUnkForRelease = nullptr;
				return S_OK;
			}

			return DV_E_FORMATETC;
		}

		UINT cfShellIDList = RegisterClipboardFormatW(CFSTR_SHELLIDLIST);
		if (pFormatEtc->cfFormat == cfShellIDList && (pFormatEtc->tymed & TYMED_HGLOBAL)) {
			HGLOBAL hg = CreateShellIDListHGlobal(_paths);
			if (!hg) return E_OUTOFMEMORY;
			pMedium->tymed = TYMED_HGLOBAL;
			pMedium->hGlobal = hg;
			pMedium->pUnkForRelease = nullptr;
			return S_OK;
		}

		return DV_E_FORMATETC;
	}

	HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC*, STGMEDIUM*) override { return E_NOTIMPL; }

	HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* pFormatEtc) override {
		if (!pFormatEtc) return E_POINTER;

		// debug log
		wchar_t dbg[512];
		swprintf_s(dbg, L"QueryGetData: cf=%s (0x%08X) tymed=0x%X lindex=%d\n",
			FormatIdToName(pFormatEtc->cfFormat).c_str(),
			pFormatEtc->cfFormat, pFormatEtc->tymed, pFormatEtc->lindex);
		OutputDebugStringW(dbg);

		if (pFormatEtc->cfFormat == CF_HDROP && (pFormatEtc->tymed & TYMED_HGLOBAL)) return S_OK;

		UINT cfInetUrlW = RegisterClipboardFormatW(CFSTR_INETURLW);
		if (pFormatEtc->cfFormat == cfInetUrlW && (pFormatEtc->tymed & TYMED_HGLOBAL)) return S_OK;

		UINT cfPref = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
		if (pFormatEtc->cfFormat == cfPref && (pFormatEtc->tymed & TYMED_HGLOBAL)) return S_OK;

		UINT cfFileDesc = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);
		if (pFormatEtc->cfFormat == cfFileDesc && (pFormatEtc->tymed & TYMED_HGLOBAL)) return S_OK;

		UINT cfFileCont = RegisterClipboardFormatW(CFSTR_FILECONTENTS);
		if (pFormatEtc->cfFormat == cfFileCont) {
			if (pFormatEtc->tymed & (TYMED_ISTREAM | TYMED_HGLOBAL)) return S_OK;
		}

		UINT cfShellIDList = RegisterClipboardFormatW(CFSTR_SHELLIDLIST); // "Shell IDList Array"
		if (pFormatEtc->cfFormat == cfShellIDList && (pFormatEtc->tymed & TYMED_HGLOBAL)) return S_OK;

		return DV_E_FORMATETC;
	}

	HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC* pFormatEtc, FORMATETC* pFormatEtcOut) override {
		if (!pFormatEtcOut) return E_POINTER;
		*pFormatEtcOut = *pFormatEtc;
		pFormatEtcOut->ptd = nullptr;
		return DATA_S_SAMEFORMATETC;
	}

	HRESULT STDMETHODCALLTYPE SetData(FORMATETC*, STGMEDIUM*, BOOL) override { return E_NOTIMPL; }

	HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC** ppEnum) override {
		if (!ppEnum) return E_POINTER;
		if (dwDirection != DATADIR_GET) return E_NOTIMPL;

		// Build the list of FORMATETC entries you support
		std::vector<FORMATETC> fmts;

		FORMATETC f = {};
		f.cfFormat = CF_HDROP; f.tymed = TYMED_HGLOBAL; f.dwAspect = DVASPECT_CONTENT; f.lindex = -1; f.ptd = nullptr;
		fmts.push_back(f);

		UINT cfInetUrlW = RegisterClipboardFormatW(CFSTR_INETURLW);
		if (cfInetUrlW) { f.cfFormat = (CLIPFORMAT)cfInetUrlW; f.tymed = TYMED_HGLOBAL; fmts.push_back(f); }

		UINT cfPref = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
		if (cfPref) { f.cfFormat = (CLIPFORMAT)cfPref; f.tymed = TYMED_HGLOBAL; fmts.push_back(f); }

		UINT cfFileDesc = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);
		if (cfFileDesc) { f.cfFormat = (CLIPFORMAT)cfFileDesc; f.tymed = TYMED_HGLOBAL; fmts.push_back(f); }

		UINT cfFileCont = RegisterClipboardFormatW(CFSTR_FILECONTENTS);
		if (cfFileCont) { f.cfFormat = (CLIPFORMAT)cfFileCont; f.tymed = (TYMED_ISTREAM | TYMED_HGLOBAL); fmts.push_back(f); }

		UINT cfShellIDList = RegisterClipboardFormatW(CFSTR_SHELLIDLIST);
		if (cfShellIDList) { f.cfFormat = (CLIPFORMAT)cfShellIDList; f.tymed = TYMED_HGLOBAL; fmts.push_back(f); }

		// allocate array and create standard enumerator
		return SHCreateStdEnumFmtEtc((UINT)fmts.size(), fmts.empty() ? nullptr : &fmts[0], ppEnum);
	}

	HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override { return OLE_E_ADVISENOTSUPPORTED; }
	HRESULT STDMETHODCALLTYPE DUnadvise(DWORD) override { return OLE_E_ADVISENOTSUPPORTED; }
	HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA**) override { return OLE_E_ADVISENOTSUPPORTED; }

private:
	// Helper: create FILEGROUPDESCRIPTORW HGLOBAL for _paths
	HGLOBAL CreateFileGroupDescriptorHGlobal() {
		size_t n = _paths.size();
		SIZE_T bytes = sizeof(FILEGROUPDESCRIPTORW) + (n ? (n - 1) * sizeof(FILEDESCRIPTORW) : 0);
		HGLOBAL hg = GlobalAlloc(GHND | GMEM_SHARE, bytes);
		if (!hg) return NULL;
		FILEGROUPDESCRIPTORW* pgd = (FILEGROUPDESCRIPTORW*)GlobalLock(hg);
		ZeroMemory(pgd, bytes);
		pgd->cItems = (UINT)n;
		for (size_t i = 0; i < n; ++i) {
			FILEDESCRIPTORW& fd = pgd->fgd[i];
			ZeroMemory(&fd, sizeof(fd));
			// filename
			std::wstring name;
			size_t pos = _paths[i].find_last_of(L"\\/");
			name = (pos == std::wstring::npos) ? _paths[i] : _paths[i].substr(pos + 1);
			wcsncpy_s(fd.cFileName, name.c_str(), ARRAYSIZE(fd.cFileName)-1);

			// try to fill size, attributes and time
			WIN32_FILE_ATTRIBUTE_DATA fad;
			if (GetFileAttributesExW(_paths[i].c_str(), GetFileExInfoStandard, &fad)) {
				fd.dwFlags = FD_FILESIZE | FD_ATTRIBUTES | FD_WRITESTIME;
				ULONGLONG size = ((ULONGLONG)fad.nFileSizeHigh << 32) | fad.nFileSizeLow;
				fd.nFileSizeLow = (DWORD)(size & 0xFFFFFFFF);
				fd.nFileSizeHigh = (DWORD)(size >> 32);
				fd.dwFileAttributes = fad.dwFileAttributes;
				fd.ftLastWriteTime = fad.ftLastWriteTime;
			}
			else {
				fd.dwFlags = 0;
			}
		}
		GlobalUnlock(hg);
		return hg;
	}
};

// IDropTarget implementation to accept .m3u files
class PlaylistDnDTarget : public IDropTarget {
	LONG _ref;
	HWND _hwnd;
public:
	PlaylistDnDTarget(HWND hwnd) : _ref(1), _hwnd(hwnd) {}
	~PlaylistDnDTarget() {}

	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override {
		if (!ppvObject) return E_POINTER;
		if (riid == IID_IUnknown || riid == IID_IDropTarget) {
			*ppvObject = static_cast<IDropTarget*>(this);
			AddRef();
			return S_OK;
		}
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}
	ULONG STDMETHODCALLTYPE AddRef(void) override { return InterlockedIncrement(&_ref); }
	ULONG STDMETHODCALLTYPE Release(void) override {
		ULONG c = InterlockedDecrement(&_ref);
		if (c == 0) delete this;
		return c;
	}

	HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) override {
		UNREFERENCED_PARAMETER(grfKeyState);
		UNREFERENCED_PARAMETER(pt);
		if (!pDataObj || !pdwEffect) return E_POINTER;
		FORMATETC fmt = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		if (pDataObj->QueryGetData(&fmt) == S_OK) {
			*pdwEffect = DROPEFFECT_COPY;
			return S_OK;
		}
		*pdwEffect = DROPEFFECT_NONE;
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE DragOver(DWORD, POINTL, DWORD* pdwEffect) override {
		if (!pdwEffect) return E_POINTER;
		*pdwEffect = DROPEFFECT_COPY;
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE DragLeave(void) override { return S_OK; }

	HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObj, DWORD, POINTL pt, DWORD* pdwEffect) override {
		if (!pDataObj || !pdwEffect) return E_POINTER;
		FORMATETC fmt = { CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		STGMEDIUM stg = {};
		if (SUCCEEDED(pDataObj->GetData(&fmt, &stg))) {
			HDROP hDrop = (HDROP)GlobalLock(stg.hGlobal);
			if (hDrop) {
				UINT count = DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0);
				auto* collected = new std::vector<std::wstring>();
				for (UINT i = 0; i < count; ++i) {
					wchar_t path[MAX_PATH];
					if (0 == DragQueryFileW(hDrop, i, path, ARRAYSIZE(path))) continue;
					std::wstring p(path);

					DWORD attr = GetFileAttributesW(p.c_str());
					if (attr != INVALID_FILE_ATTRIBUTES && (attr & FILE_ATTRIBUTE_DIRECTORY)) {
						std::vector<std::wstring> filter;
						EnumerateDirectoryFiles(p, *collected, filter);
						continue;
					}

					std::wstring ext;
					size_t pos = p.find_last_of(L'.');
					if (pos != std::wstring::npos) ext = p.substr(pos);
					std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
					if (ext == L".m3u" || ext == L".m3u8") {
						std::vector<std::wstring> lines = ReadM3ULines(p);
						size_t bpos = p.find_last_of(L"\\/");
						std::wstring baseDir = (bpos == std::wstring::npos) ? L"" : p.substr(0, bpos);
						for (auto& ln : lines) {
							if (ln.empty() || ln[0] == L'#') continue;
							std::wstring resolved = ResolvePath(baseDir, ln);
							if (GetFileAttributesW(resolved.c_str()) != INVALID_FILE_ATTRIBUTES) {
								collected->push_back(resolved);
							}
						}
						continue;
					}

					if (GetFileAttributesW(p.c_str()) != INVALID_FILE_ATTRIBUTES) {
						collected->push_back(p);
					}
				}

				GlobalUnlock(stg.hGlobal);
				// inside PlaylistDnDTarget::Drop, after building 'collected' and before ReleaseStgMedium
				// compute insertion index relative to the listview (g_hList)
				int insertIndex = -1;
				if (IsWindow(g_hList)) {
					// pt is POINTL (screen coords) passed to Drop; convert to client coords of the listview
					POINT p = { (LONG)pt.x, (LONG)pt.y };
					ScreenToClient(g_hList, &p);

					LVHITTESTINFO hti = {};
					hti.pt = p;
					int hit = ListView_HitTest(g_hList, &hti);
					if (hit >= 0) {
						// decide before/after based on vertical midpoint of the item rect
						RECT rcItem;
						if (ListView_GetItemRect(g_hList, hit, &rcItem, LVIR_BOUNDS)) {
							int mid = (rcItem.top + rcItem.bottom) / 2;
							insertIndex = (p.y < mid) ? hit : (hit + 1);
						}
						else {
							insertIndex = hit + 1;
						}
					}
					else {
						// no item under pointer -> insert at end
						insertIndex = ListView_GetItemCount(g_hList);
					}
				}
				else {
					insertIndex = -1; // fallback: let WM handler append
				}

				// Post insertion index in wParam and pointer in lParam
				PostMessageW(_hwnd, WM_APP + 2, (WPARAM)insertIndex, (LPARAM)collected);
				*pdwEffect = DROPEFFECT_COPY;
				ReleaseStgMedium(&stg);
				return S_OK;
			}
			ReleaseStgMedium(&stg);
		}
		*pdwEffect = DROPEFFECT_NONE;
		return S_OK;
	}
};

class SimpleDropSource : public IDropSource {
	LONG _ref;
public:
	SimpleDropSource() : _ref(1) {}
	~SimpleDropSource() {}
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override {
		if (!ppvObject) return E_POINTER;
		if (riid == IID_IUnknown || riid == IID_IDropSource) {
			*ppvObject = static_cast<IDropSource*>(this);
			AddRef();
			return S_OK;
		}
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}
	ULONG STDMETHODCALLTYPE AddRef(void) override { return InterlockedIncrement(&_ref); }
	ULONG STDMETHODCALLTYPE Release(void) override {
		ULONG c = InterlockedDecrement(&_ref);
		if (c == 0) delete this;
		return c;
	}
	HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState) override {
		if (fEscapePressed) return DRAGDROP_S_CANCEL;
		if (!(grfKeyState & MK_LBUTTON)) return DRAGDROP_S_DROP;
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD) override {
		return DRAGDROP_S_USEDEFAULTCURSORS;
	}
};

// retourne durée en millisecondes si OK, sinon 0
inline ULONGLONG GetMediaDurationMs(const std::wstring& path) {
    IPropertyStore* pps = nullptr;
    HRESULT hr = SHGetPropertyStoreFromParsingName(path.c_str(), NULL, GPS_DEFAULT, IID_PPV_ARGS(&pps));
    if (FAILED(hr) || !pps) return 0;

    PROPVARIANT pv;
    PropVariantInit(&pv);
    hr = pps->GetValue(PKEY_Media_Duration, &pv);
    ULONGLONG durationMs = 0;
    if (SUCCEEDED(hr)) {
        // PKEY_Media_Duration is typically VT_UI8 and expressed in 100-ns units
        if (pv.vt == VT_UI8) {
            ULONGLONG hundredNs = pv.uhVal.QuadPart;
            // convert 100-ns units to milliseconds
            durationMs = (hundredNs + 5000) / 10000;
        } else if (pv.vt == VT_I8) {
            LONGLONG v = pv.hVal.QuadPart;
            if (v > 0) durationMs = (ULONGLONG)((v + 5000) / 10000);
        }
    }
    PropVariantClear(&pv);
    pps->Release();
    return durationMs;
}

inline std::wstring FormatDurationMs(ULONGLONG ms) {
    if (ms == 0) return L"";
    ULONGLONG s = ms / 1000;
    ULONGLONG mins = s / 60;
    ULONGLONG secs = s % 60;
    wchar_t buf[64];
    swprintf_s(buf, L"%llu:%02llu", mins, secs);
    return std::wstring(buf);
}
