#pragma once

#include "resource.h"
#include "Utils.h"
#include "M3UUtils.h"
#include <random>

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
COLORREF g_clrCtrl = RGB(36, 36, 48);// RGB(18, 24, 48);// RGB(28, 34, 64);      // control background (slightly lighter)
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

class FileDropDataObject : public IDataObject {
	LONG _ref;
	HGLOBAL _hDrop;
public:
	FileDropDataObject(const std::vector<std::wstring>& paths) : _ref(1), _hDrop(NULL) {
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

	HRESULT STDMETHODCALLTYPE GetData(FORMATETC* pFormatEtc, STGMEDIUM* pMedium) override {
		if (!pFormatEtc || !pMedium) return E_POINTER;
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
		return DV_E_FORMATETC;
	}

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

	HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC*, STGMEDIUM*) override { return E_NOTIMPL; }
	HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* pFormatEtc) override {
		if (!pFormatEtc) return E_POINTER;
		if ((pFormatEtc->cfFormat == CF_HDROP) && (pFormatEtc->tymed & TYMED_HGLOBAL)) return S_OK;
		return DV_E_FORMATETC;
	}
	HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC*, FORMATETC* pOut) override {
		if (!pOut) return E_POINTER;
		pOut->ptd = nullptr;
		return E_NOTIMPL;
	}
	HRESULT STDMETHODCALLTYPE SetData(FORMATETC*, STGMEDIUM*, BOOL) override { return E_NOTIMPL; }
	HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD, IEnumFORMATETC**) override { return E_NOTIMPL; }
	HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override { return OLE_E_ADVISENOTSUPPORTED; }
	HRESULT STDMETHODCALLTYPE DUnadvise(DWORD) override { return OLE_E_ADVISENOTSUPPORTED; }
	HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA**) override { return OLE_E_ADVISENOTSUPPORTED; }
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
