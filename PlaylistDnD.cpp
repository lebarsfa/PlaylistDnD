#include "framework.h"
#include "PlaylistDnD.h"

void FreeAllListViewItems(HWND hList) {
	if (!IsWindow(hList)) return;
	int count = ListView_GetItemCount(hList);
	for (int i = 0; i < count; ++i) {
		LVITEMW it = {};
		it.mask = LVIF_PARAM;
		it.iItem = i;
		if (ListView_GetItem(hList, &it)) {
			std::wstring* p = reinterpret_cast<std::wstring*>(it.lParam);
			delete p;
		}
	}
	ListView_DeleteAllItems(hList);
}

void UpdateListViewColumns(HWND hList)
{
	if (!IsWindow(hList)) return;
	if (!g_colUserSized[0]) ListView_SetColumnWidth(hList, 0, LVSCW_AUTOSIZE_USEHEADER);
	if (!g_colUserSized[1]) ListView_SetColumnWidth(hList, 1, LVSCW_AUTOSIZE_USEHEADER);
	if (!g_colUserSized[2]) ListView_SetColumnWidth(hList, 2, LVSCW_AUTOSIZE_USEHEADER);
	if (!g_colUserSized[3]) ListView_SetColumnWidth(hList, 3, LVSCW_AUTOSIZE_USEHEADER);
}

void ShuffleListView(HWND hList)
{
	if (!IsWindow(hList)) return;

	int count = ListView_GetItemCount(hList);
	if (count <= 1) return;

	// 1) Collect pointers (lParam) in order
	std::vector<std::wstring*> items;
	items.reserve(count);
	for (int i = 0; i < count; ++i) {
		LVITEMW it = {};
		it.mask = LVIF_PARAM;
		it.iItem = i;
		if (ListView_GetItem(hList, &it)) {
			std::wstring* p = reinterpret_cast<std::wstring*>(it.lParam);
			items.push_back(p);
		}
		else {
			items.push_back(nullptr);
		}
	}

	// 2) Save selected pointers so we can reselect them after shuffle
	std::vector<std::wstring*> selectedPtrs;
	int idx = -1;
	while (true) {
		idx = ListView_GetNextItem(hList, idx, LVNI_SELECTED);
		if (idx == -1) break;
		LVITEMW it = {};
		it.mask = LVIF_PARAM;
		it.iItem = idx;
		if (ListView_GetItem(hList, &it)) {
			selectedPtrs.push_back(reinterpret_cast<std::wstring*>(it.lParam));
		}
	}

	// 3) Shuffle using random_device + mt19937
	std::random_device rd;
	std::mt19937 g(rd());
	std::shuffle(items.begin(), items.end(), g);

	// 4) Remove all items from the ListView (do NOT free lParam pointers here)
	ListView_DeleteAllItems(hList);

	// 5) Reinsert items in shuffled order, reusing the same lParam pointers
	for (size_t i = 0; i < items.size(); ++i) {
		std::wstring* pFull = items[i];
		std::wstring nameOnlyNoExt = L"";
		std::wstring extension = L"";
		std::wstring length = L"";
		std::wstring directory = L"";
		if (pFull && !pFull->empty()) {
			std::wstring filename = *pFull;
			size_t ppos = filename.find_last_of(L"\\/");
			directory = (ppos == std::wstring::npos) ? L"" : filename.substr(0, ppos);
			std::wstring nameOnly = (ppos == std::wstring::npos) ? filename : filename.substr(ppos + 1);
			size_t dot = nameOnly.find_last_of(L'.');
			nameOnlyNoExt = (dot == std::wstring::npos) ? nameOnly : nameOnly.substr(0, dot);
			extension = (dot == std::wstring::npos) ? L"" : nameOnly.substr(dot);
			length = FormatDurationMs(GetMediaDurationMs(filename));
		}

		LVITEMW item = {};
		item.mask = LVIF_TEXT | LVIF_PARAM;
		item.iItem = ListView_GetItemCount(hList);
		item.iSubItem = 0;
		item.pszText = const_cast<LPWSTR>(nameOnlyNoExt.c_str());
		item.lParam = (LPARAM)pFull;
		int newIdx = ListView_InsertItem(hList, &item);
		if (newIdx >= 0) {
			ListView_SetItemText(hList, newIdx, 1, const_cast<LPWSTR>(extension.c_str()));
			ListView_SetItemText(hList, newIdx, 2, const_cast<LPWSTR>(length.c_str()));
			ListView_SetItemText(hList, newIdx, 3, const_cast<LPWSTR>(directory.c_str()));
		}
	}

	// 6) Restore selection: find items whose lParam is in selectedPtrs
	if (!selectedPtrs.empty()) {
		for (int i = 0; i < ListView_GetItemCount(hList); ++i) {
			LVITEMW it = {};
			it.mask = LVIF_PARAM;
			it.iItem = i;
			if (ListView_GetItem(hList, &it)) {
				std::wstring* p = reinterpret_cast<std::wstring*>(it.lParam);
				if (p) {
					// linear search in selectedPtrs (small lists are fine)
					for (auto sp : selectedPtrs) {
						if (sp == p) {
							ListView_SetItemState(hList, i, LVIS_SELECTED, LVIS_SELECTED);
							break;
						}
					}
				}
			}
		}
		// set focus to first selected item
		int firstSel = ListView_GetNextItem(hList, -1, LVNI_SELECTED);
		if (firstSel != -1) ListView_SetItemState(hList, firstSel, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
	}

	// 7) Update columns (respects g_colUserSized[]) and repaint
	UpdateListViewColumns(hList);
	InvalidateRect(hList, NULL, TRUE);
	UpdateWindow(hList);
}

void LoadM3UToList(HWND hList, const std::wstring& playlistPath)
{
	if (!IsWindow(hList)) return;
	FreeAllListViewItems(hList);

	std::vector<std::wstring> lines = ReadM3ULines(playlistPath);
	if (lines.empty()) {
		OutputDebugStringW(L"LoadM3UToList: no lines read.\n");
		return;
	}

	std::wstring baseDir;
	size_t pos = playlistPath.find_last_of(L"\\/");
	if (pos != std::wstring::npos) baseDir = playlistPath.substr(0, pos);

	//wchar_t prevDir[MAX_PATH] = { 0 };
	//DWORD prevLen = GetCurrentDirectoryW(MAX_PATH, prevDir);

	if (!baseDir.empty()) {
		if (!SetCurrentDirectoryW(baseDir.c_str())) {
			DWORD err = GetLastError();
			wchar_t buf[256];
			swprintf_s(buf, L"SetCurrentDirectoryW failed for '%s' (err=%u)\n", baseDir.c_str(), err);
			OutputDebugStringW(buf);
		}
		else {
			OutputDebugStringW(L"Current directory set to playlist folder.\n");
		}
	}

	int added = 0;
	for (const auto& raw : lines) {
		if (raw.empty()) continue;
		if (raw[0] == L'#') continue;

		std::wstring resolved = ResolvePath(baseDir, raw);
		if (GetFileAttributesW(resolved.c_str()) == INVALID_FILE_ATTRIBUTES) {
			OutputDebugStringW(L"File not found: ");
			OutputDebugStringW(resolved.c_str());
			OutputDebugStringW(L"\n");
		}

		std::wstring filename = resolved;
		size_t ppos = filename.find_last_of(L"\\/");
		std::wstring directory = (ppos == std::wstring::npos) ? L"" : filename.substr(0, ppos);
		std::wstring nameOnly = (ppos == std::wstring::npos) ? filename : filename.substr(ppos + 1);
		size_t dot = nameOnly.find_last_of(L'.');
		std::wstring nameOnlyNoExt = (dot == std::wstring::npos) ? nameOnly : nameOnly.substr(0, dot);
		std::wstring extension = (dot == std::wstring::npos) ? L"" : nameOnly.substr(dot);
		std::wstring length = FormatDurationMs(GetMediaDurationMs(filename));

		LVITEMW item = {};
		item.mask = LVIF_TEXT | LVIF_PARAM;
		item.iItem = ListView_GetItemCount(hList);
		item.iSubItem = 0;
		item.pszText = const_cast<LPWSTR>(nameOnlyNoExt.c_str());
		item.lParam = (LPARAM)new std::wstring(resolved);
		int idx = ListView_InsertItem(hList, &item);
		if (idx >= 0) {
			ListView_SetItemText(hList, idx, 1, const_cast<LPWSTR>(extension.c_str()));
			ListView_SetItemText(hList, idx, 2, const_cast<LPWSTR>(length.c_str()));
			ListView_SetItemText(hList, idx, 3, const_cast<LPWSTR>(directory.c_str()));
			++added;
		}
	}

	if (ListView_GetItemCount(hList) > 0) ListView_SetItemState(hList, 0, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
}

void InsertPathToListView(HWND hList, const std::wstring& fullPath)
{
	if (!IsWindow(hList)) return;
	std::wstring filename = fullPath;
	size_t pos = filename.find_last_of(L"\\/");
	std::wstring directory = (pos == std::wstring::npos) ? L"" : filename.substr(0, pos);
	std::wstring nameOnly = (pos == std::wstring::npos) ? filename : filename.substr(pos + 1);
	size_t dot = nameOnly.find_last_of(L'.');
	std::wstring nameOnlyNoExt = (dot == std::wstring::npos) ? nameOnly : nameOnly.substr(0, dot);
	std::wstring extension = (dot == std::wstring::npos) ? L"" : nameOnly.substr(dot);
	std::wstring length = FormatDurationMs(GetMediaDurationMs(filename));

	LVITEMW item = {};
	item.mask = LVIF_TEXT | LVIF_PARAM;
	item.iItem = ListView_GetItemCount(hList);
	item.iSubItem = 0;
	item.pszText = const_cast<LPWSTR>(nameOnlyNoExt.c_str());
	item.lParam = (LPARAM)new std::wstring(fullPath);
	int idx = ListView_InsertItem(hList, &item);
	if (idx >= 0) {
		ListView_SetItemText(hList, idx, 1, const_cast<LPWSTR>(extension.c_str()));
		ListView_SetItemText(hList, idx, 2, const_cast<LPWSTR>(length.c_str()));
		ListView_SetItemText(hList, idx, 3, const_cast<LPWSTR>(directory.c_str()));
	}
}

std::wstring GetFullPathFromItem(HWND hList, int index)
{
	LVITEMW it = {};
	it.mask = LVIF_PARAM;
	it.iItem = index;
	it.iSubItem = 0;
	if (ListView_GetItem(hList, &it)) {
		std::wstring* p = reinterpret_cast<std::wstring*>(it.lParam);
		return p ? *p : std::wstring();
	}
	return std::wstring();
}

std::vector<int> GetSelectedListViewIndices(HWND hList)
{
	std::vector<int> out;
	int idx = -1;
	while (true) {
		idx = ListView_GetNextItem(hList, idx, LVNI_SELECTED);
		if (idx == -1) break;
		out.push_back(idx);
	}
	return out;
}

void ClearListViewSelection(HWND hList)
{
	if (!IsWindow(hList)) return;
	auto indices = GetSelectedListViewIndices(hList);
	if (indices.empty()) return;
	SendMessageW(hList, WM_SETREDRAW, FALSE, 0);
	for (int idx : indices) {
		ListView_SetItemState(hList, idx, 0, LVIS_SELECTED);
	}
	SendMessageW(hList, WM_SETREDRAW, TRUE, 0);
	InvalidateRect(hList, NULL, TRUE);
	UpdateWindow(hList);
}

void DeleteSelectedListViewItems(HWND hList)
{
	if (!IsWindow(hList)) return;
	std::vector<int> indices = GetSelectedListViewIndices(hList);
	if (indices.empty()) return;
	std::sort(indices.begin(), indices.end(), std::greater<int>());
	for (int idx : indices) {
		LVITEMW it = {};
		it.mask = LVIF_PARAM;
		it.iItem = idx;
		if (ListView_GetItem(hList, &it)) {
			std::wstring* p = reinterpret_cast<std::wstring*>(it.lParam);
			delete p;
		}
		ListView_DeleteItem(hList, idx);
	}
	UpdateListViewColumns(hList);
}

void StartDragFromListMulti(HWND hList)
{
	if (!IsWindow(hList)) return;

	std::vector<int> indices = GetSelectedListViewIndices(hList);
	if (indices.empty()) {
		int focused = ListView_GetNextItem(hList, -1, LVNI_FOCUSED);
		if (focused != -1) indices.push_back(focused);
	}
	if (indices.empty()) return;

	std::vector<std::wstring> paths;
	for (int idx : indices) {
		std::wstring p = GetFullPathFromItem(hList, idx);
		if (!p.empty() && GetFileAttributesW(p.c_str()) != INVALID_FILE_ATTRIBUTES) {
			paths.push_back(p);
		}
		else {
			OutputDebugStringW(L"StartDragFromListMulti: skipping non-existent: ");
			OutputDebugStringW(p.c_str());
			OutputDebugStringW(L"\n");
		}
	}
	if (paths.empty()) return;

	for (auto& p : paths) {
		OutputDebugStringW(L"StartDragFromListMulti: will drop: ");
		OutputDebugStringW(p.c_str());
		OutputDebugStringW(L"\n");
	}

	FileDropDataObject* pData = new FileDropDataObject(paths);
	pData->AddRef();

	SimpleDropSource* pDropSrc = new SimpleDropSource();

	DWORD dwEffect = 0;
	HRESULT hr = DoDragDrop(pData, pDropSrc, DROPEFFECT_COPY, &dwEffect);

	if (hr == DRAGDROP_S_DROP) {
		OutputDebugStringW(L"DoDragDrop: DRAGDROP_S_DROP\n");
	}
	else if (hr == DRAGDROP_S_CANCEL) {
		OutputDebugStringW(L"DoDragDrop: DRAGDROP_S_CANCEL\n");
	}
	else {
		wchar_t buf2[128];
		swprintf_s(buf2, L"DoDragDrop: hr=0x%08X dwEffect=0x%08X\n", hr, dwEffect);
		OutputDebugStringW(buf2);
	}

	pDropSrc->Release();
	pData->Release();
}

LRESULT CALLBACK ListboxSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
	UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	UNREFERENCED_PARAMETER(dwRefData);

	switch (msg) {
	case WM_LBUTTONDOWN: {
		POINT ptClient = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
		LVHITTESTINFO hti = {};
		hti.pt = ptClient;
		int idx = ListView_HitTest(hwnd, &hti);
		bool onItem = (idx != -1) && (hti.flags & LVHT_ONITEM);

		bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
		bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

		if (onItem) {
			UINT state = ListView_GetItemState(hwnd, idx, LVIS_SELECTED);
			if ((state & LVIS_SELECTED) && !ctrl && !shift) {
				SetFocus(hwnd);
				s_dt.tracking = true;
				s_dt.downPt = ptClient;
				SetCapture(hwnd);
				return 0;
			}
		}
		return DefSubclassProc(hwnd, msg, wParam, lParam);
	}

	case WM_MOUSEMOVE: {
		if (s_dt.tracking) {
			POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
			int dx = abs(pt.x - s_dt.downPt.x);
			int dy = abs(pt.y - s_dt.downPt.y);
			int sx = GetSystemMetrics(SM_CXDRAG);
			int sy = GetSystemMetrics(SM_CYDRAG);
			if (dx >= sx || dy >= sy) {
				StartDragFromListMulti(hwnd);
				s_dt.tracking = false;
				ReleaseCapture();
				return 0;
			}
		}
		break;
	}

	case WM_LBUTTONUP: {
		if (s_dt.tracking) {
			s_dt.tracking = false;
			ReleaseCapture();
			return 0;
		}
		break;
	}

	case WM_KEYDOWN: {
		bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
		if (ctrl && (wParam == 'A' || wParam == 'a')) {
			int count = ListView_GetItemCount(hwnd);
			for (int i = 0; i < count; ++i) {
				ListView_SetItemState(hwnd, i, LVIS_SELECTED, LVIS_SELECTED);
			}
			if (count > 0) ListView_SetItemState(hwnd, 0, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
			return 0;
		}
		if (wParam == VK_DELETE) {
			DeleteSelectedListViewItems(hwnd);
			return 0;
		}
		break;
	}

	case WM_RBUTTONDOWN: {
		ClearListViewSelection(hwnd);
		break;
	}

	case WM_NCDESTROY: {
		RemoveWindowSubclass(hwnd, ListboxSubclassProc, uIdSubclass);
		break;
	}
	}

	return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
	case WM_CREATE: {
		INITCOMMONCONTROLSEX icex = {};
		icex.dwSize = sizeof(icex);
		icex.dwICC = ICC_LISTVIEW_CLASSES;
		InitCommonControlsEx(&icex);

		hLabelFont = CreateUiFont(hWnd, L"Segoe UI", 10); // 10 pt Segoe UI
		if (!hLabelFont) hLabelFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

		g_hLabel = CreateWindowExW(0, L"STATIC",
			L"Drag and drop files, folders or .m3u playlists below, then select and drag and drop the desired items to your media player...",
			WS_CHILD | WS_VISIBLE | SS_LEFT,
			10, 10, 100, 20, hWnd, (HMENU)2001, g_hInst, NULL);

		SendMessageW(g_hLabel, WM_SETFONT, (WPARAM)hLabelFont, TRUE);

		g_hList = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, NULL,
			WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS |
			WS_VSCROLL | WS_HSCROLL,
			10, 40, 100, 100, hWnd, (HMENU)1001, g_hInst, NULL);

		ListView_SetExtendedListViewStyle(g_hList,
			LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_LABELTIP);

		LVCOLUMNW col = {};
		col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

		col.cx = 200;;
		col.pszText = const_cast<LPWSTR>(L"Filename");
		ListView_InsertColumn(g_hList, 0, &col);

		col.cx = 25;
		col.pszText = const_cast<LPWSTR>(L"Extension");
		ListView_InsertColumn(g_hList, 1, &col);

		col.cx = 25;
		col.pszText = const_cast<LPWSTR>(L"Length");
		ListView_InsertColumn(g_hList, 2, &col);

		col.cx = 200;
		col.pszText = const_cast<LPWSTR>(L"Path");
		ListView_InsertColumn(g_hList, 3, &col);

		SetWindowSubclass(g_hList, ListboxSubclassProc, 1, 0);

		g_dropTarget = new PlaylistDnDTarget(hWnd);
		RegisterDragDrop(hWnd, g_dropTarget);

		g_hbrBg = CreateSolidBrush(g_clrBg);
		g_hbrCtrl = CreateSolidBrush(g_clrCtrl);

		// Keep ListView background set but custom draw will override subitems
		ListView_SetBkColor(g_hList, g_clrCtrl);
		ListView_SetTextColor(g_hList, g_clrText);
		ListView_SetTextBkColor(g_hList, g_clrCtrl);

		g_hHeader = ListView_GetHeader(g_hList);
		if (g_hHeader) {
			SendMessageW(g_hHeader, HDM_SETBKCOLOR, 0, (LPARAM)g_clrCtrl);
			InvalidateRect(g_hHeader, NULL, TRUE);
		}

		g_colUserSized[0] = g_colUserSized[1] = g_colUserSized[2] = g_colUserSized[3] = false;

		SendMessageW(g_hList, WM_SETFONT, (WPARAM)hLabelFont, TRUE);

		g_hShuffleBtn = CreateWindowExW(0, L"BUTTON", L"Shuffle",
			WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
			0, 0, 80, 26, hWnd, (HMENU)ID_BTN_SHUFFLE, g_hInst, NULL);

		SendMessageW(g_hShuffleBtn, WM_SETFONT, (WPARAM)hLabelFont, TRUE);

		RECT rc;
		GetClientRect(hWnd, &rc);
		SendMessageW(hWnd, WM_SIZE, 0, MAKELPARAM(rc.right - rc.left, rc.bottom - rc.top));
		break;
	}

	case WM_DESTROY: {
		if (g_dropTarget) {
			RevokeDragDrop(hWnd);
			g_dropTarget->Release();
			g_dropTarget = nullptr;
		}
		FreeAllListViewItems(g_hList);
		if (g_hbrBg) { DeleteObject(g_hbrBg); g_hbrBg = NULL; }
		if (g_hbrCtrl) { DeleteObject(g_hbrCtrl); g_hbrCtrl = NULL; }
		if (g_hbrScrollBg) { DeleteObject(g_hbrScrollBg); g_hbrScrollBg = NULL; }
		if (g_hbrScrollThumb) { DeleteObject(g_hbrScrollThumb); g_hbrScrollThumb = NULL; }
		if (hLabelFont) DeleteObject(hLabelFont);
		PostQuitMessage(0);
		break;
	}

	case WM_ERASEBKGND: {
		HDC hdc = (HDC)wParam;
		RECT rc;
		GetClientRect(hWnd, &rc);
		FillRect(hdc, &rc, g_hbrBg);
		return 1;
	}

	case WM_CTLCOLORSTATIC:
	case WM_CTLCOLOREDIT:
	case WM_CTLCOLORLISTBOX: {
		HDC hdc = (HDC)wParam;
		SetTextColor(hdc, g_clrText);
		SetBkMode(hdc, OPAQUE);
		SetBkColor(hdc, g_clrCtrl);
		return (LRESULT)g_hbrCtrl;
	}

	case WM_SIZE: {
		int cx = LOWORD(lParam);
		int cy = HIWORD(lParam);

		const int margin = 10;
		const int labelHeight = 40;
		const int btnW = 90;
		const int btnH = 26;

		if (IsWindow(g_hLabel)) {
			SetWindowPos(g_hLabel, NULL,
				margin, margin,
				cx - btnW - margin * 3, labelHeight,
				SWP_NOZORDER | SWP_SHOWWINDOW);
		}

		if (IsWindow(g_hList)) {
			int listTop = margin + labelHeight + 6;
			int listHeight = max(0, cy - listTop - margin);

			SetWindowPos(g_hList, NULL,
				margin, listTop,
				max(0, cx - margin * 2), listHeight,
				SWP_NOZORDER | SWP_SHOWWINDOW);

			int clientWidth = max(0, cx - margin * 6);
			int col0 = (int)(clientWidth * 0.45);
			int col1 = (int)(clientWidth * 0.05);
			int col2 = (int)(clientWidth * 0.05);
			int col3 = (int)(clientWidth * 0.45);// clientWidth - col0 - col1 - col2;

			const int minCol = 40;
			if (col0 < minCol) col0 = minCol;
			if (col1 < minCol) col1 = minCol;
			if (col2 < minCol) col2 = minCol;
			if (col3 < minCol) col3 = minCol;

			if (!g_colUserSized[0]) ListView_SetColumnWidth(g_hList, 0, col0);
			if (!g_colUserSized[1]) ListView_SetColumnWidth(g_hList, 1, col1);
			if (!g_colUserSized[2]) ListView_SetColumnWidth(g_hList, 2, col2);
			if (!g_colUserSized[3]) ListView_SetColumnWidth(g_hList, 3, col3);
		}

		if (IsWindow(g_hShuffleBtn)) {
			int bx = max(margin, cx - margin - btnW);
			int by = margin;
			SetWindowPos(g_hShuffleBtn, NULL, bx, by, btnW, btnH, SWP_NOZORDER | SWP_SHOWWINDOW);
		}

		InvalidateRect(hWnd, NULL, TRUE);
		UpdateWindow(hWnd);
		break;
	}

	case WM_APP + 1: {
		std::wstring* p = (std::wstring*)lParam;
		if (p) {
			LoadM3UToList(g_hList, *p);
			delete p;
			UpdateListViewColumns(g_hList);
			InvalidateRect(g_hList, NULL, TRUE);
			UpdateWindow(g_hList);
		}
		break;
	}

	case WM_APP + 2: {
		auto* vec = reinterpret_cast<std::vector<std::wstring>*>(lParam);
		int insertIndex = (int)wParam; // -1 means append at end
		if (vec) {
			// disable redraw for smoother insertion
			SendMessageW(g_hList, WM_SETREDRAW, FALSE, 0);

			// clamp insertIndex
			int count = ListView_GetItemCount(g_hList);
			if (insertIndex < 0 || insertIndex > count) insertIndex = count;

			// Insert each path in order at insertIndex, incrementing so order is preserved
			for (auto& p : *vec) {
				// prepare display strings
				std::wstring filename = p;
				size_t pos = filename.find_last_of(L"\\/");
				std::wstring directory = (pos == std::wstring::npos) ? L"" : filename.substr(0, pos);
				std::wstring nameOnly = (pos == std::wstring::npos) ? filename : filename.substr(pos + 1);
				size_t dot = nameOnly.find_last_of(L'.');
				std::wstring nameOnlyNoExt = (dot == std::wstring::npos) ? nameOnly : nameOnly.substr(0, dot);
				std::wstring extension = (dot == std::wstring::npos) ? L"" : nameOnly.substr(dot);
				std::wstring length = FormatDurationMs(GetMediaDurationMs(filename));

				LVITEMW item = {};
				item.mask = LVIF_TEXT | LVIF_PARAM;
				item.iItem = insertIndex; // insert at this index
				item.iSubItem = 0;
				item.pszText = const_cast<LPWSTR>(nameOnlyNoExt.c_str());
				item.lParam = (LPARAM)new std::wstring(p); // keep same ownership model as elsewhere

				// InsertItem inserts at iItem; if iItem == count it appends
				int newIdx = ListView_InsertItem(g_hList, &item);
				if (newIdx >= 0) {
					ListView_SetItemText(g_hList, newIdx, 1, const_cast<LPWSTR>(extension.c_str()));
					ListView_SetItemText(g_hList, newIdx, 2, const_cast<LPWSTR>(length.c_str()));
					ListView_SetItemText(g_hList, newIdx, 3, const_cast<LPWSTR>(directory.c_str()));
					// next insertion should go after this one
					insertIndex = newIdx + 1;
				}
				else {
					// fallback: append
					insertIndex = ListView_GetItemCount(g_hList);
				}
			}

			// cleanup
			delete vec;

			// re-enable redraw and refresh
			SendMessageW(g_hList, WM_SETREDRAW, TRUE, 0);
			UpdateListViewColumns(g_hList);
			InvalidateRect(g_hList, NULL, TRUE);
			UpdateWindow(g_hList);
		}
		break;
	}

	case WM_COMMAND: {
		int id = LOWORD(wParam);
		if (id == ID_BTN_SHUFFLE) {
			ShuffleListView(g_hList);
			return 0;
		}
		break;
	}

	case WM_NOTIFY: {
		LPNMHDR pnm = (LPNMHDR)lParam;
		if (!pnm) break;

		// Detect header end-track (user resized a column)
		if (pnm->hwndFrom == g_hHeader) {
			if (pnm->code == HDN_ENDTRACKW || pnm->code == HDN_ENDTRACKA
#ifdef HDN_ENDTRACK
				|| pnm->code == HDN_ENDTRACK
#endif
				) {
				NMHEADER* ph = (NMHEADER*)lParam;
				int col = ph->iItem;
				if (col >= 0 && col < 4) {
					g_colUserSized[col] = true;
				}
			}

			// Header custom draw (paint header background + text)
			if (pnm->code == NM_CUSTOMDRAW) {
				LPNMCUSTOMDRAW pcd = (LPNMCUSTOMDRAW)lParam;
				switch (pcd->dwDrawStage) {
				case CDDS_PREPAINT:
					return CDRF_NOTIFYITEMDRAW;
				case CDDS_ITEMPREPAINT: {
					HDC hdc = pcd->hdc;
					RECT rc = pcd->rc;
					// background
					FillRect(hdc, &rc, g_hbrCtrl);
					// text
					SetTextColor(hdc, g_clrText);
					SetBkMode(hdc, TRANSPARENT);
					HDITEMW hdi = {};
					wchar_t buf[256] = {};
					hdi.mask = HDI_TEXT;
					hdi.pszText = buf;
					hdi.cchTextMax = ARRAYSIZE(buf);
					int idx = (int)pcd->dwItemSpec;
					Header_GetItem(g_hHeader, idx, &hdi);
					DrawTextW(hdc, buf, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
					// bottom border
					HPEN hPen = CreatePen(PS_SOLID, 1, RGB(48, 56, 96));
					HPEN hOld = (HPEN)SelectObject(hdc, hPen);
					MoveToEx(hdc, rc.left, rc.bottom - 1, NULL);
					LineTo(hdc, rc.right, rc.bottom - 1);
					SelectObject(hdc, hOld);
					DeleteObject(hPen);
					return CDRF_SKIPDEFAULT; // important: prevent theme from overpainting
				}
				default: break;
				}
			}
		}

		// ListView custom draw: request item then subitem notifications and paint subitems
		if (pnm->hwndFrom == g_hList && pnm->code == NM_CUSTOMDRAW) {
			LPNMLVCUSTOMDRAW pcd = (LPNMLVCUSTOMDRAW)lParam;
			switch (pcd->nmcd.dwDrawStage) {
			case CDDS_PREPAINT:
				return CDRF_NOTIFYITEMDRAW;
			case CDDS_ITEMPREPAINT:
				return CDRF_NOTIFYSUBITEMDRAW | CDRF_NOTIFYPOSTPAINT; // demander postpaint
			case CDDS_SUBITEM | CDDS_ITEMPREPAINT: {
				// This is the subitem stage where we can set per-column background
				int sub = pcd->iSubItem;
				//bool selected = (pcd->nmcd.uItemState & CDIS_SELECTED) != 0;
				//if (selected) {
				//    pcd->clrTextBk = g_clrSelBg;
				//    pcd->clrText = g_clrSelText;
				//} else {
					// --- ALTERNATE BY COLUMN INDEX ---
				if ((sub % 2) == 0) pcd->clrTextBk = g_clrEvenCol;
				else pcd->clrTextBk = g_clrOddCol;
				pcd->clrText = g_clrText;
				//}
				return CDRF_NEWFONT;
			}
			case CDDS_ITEMPOSTPAINT: {
				// dessiner la ligne horizontale et les séparateurs verticaux pour la ligne iItem
				HDC hdc = pcd->nmcd.hdc;
				int iItem = (int)pcd->nmcd.dwItemSpec;

				// 1) horizontal line (bas de la ligne)
				RECT rcItem;
				rcItem.left = 0;
				rcItem.top = 0;
				// obtenir rectangle de l'item complet
				if (ListView_GetItemRect(g_hList, iItem, &rcItem, LVIR_BOUNDS)) {
					HPEN hPen = CreatePen(PS_SOLID, 1, gridColor);
					HPEN hOld = (HPEN)SelectObject(hdc, hPen);
					MoveToEx(hdc, rcItem.left, rcItem.bottom - 1, NULL);
					LineTo(hdc, rcItem.right, rcItem.bottom - 1);
					SelectObject(hdc, hOld);
					DeleteObject(hPen);
				}

				// 2) vertical separators : calculer positions de colonnes
				int colCount = 4; // ou ListView_GetColumnCount si vous gérez dynamiquement
				int x = 0;
				// obtenir la position X du début (client left)
				// itérer sur colonnes et additionner ListView_GetColumnWidth
				for (int c = 0; c < colCount; ++c) {
					int w = ListView_GetColumnWidth(g_hList, c);
					x += w;
					// dessiner ligne verticale à x (si vous voulez)
					HPEN hPenV = CreatePen(PS_SOLID, 1, gridColorVertical);
					HPEN hOldV = (HPEN)SelectObject(hdc, hPenV);
					MoveToEx(hdc, x - 1, rcItem.top, NULL);
					LineTo(hdc, x - 1, rcItem.bottom);
					SelectObject(hdc, hOldV);
					DeleteObject(hPenV);
				}
				return CDRF_DODEFAULT;
			}
			default:
				break;
			}
		}

		// Sorting on column click
		if (pnm->hwndFrom == g_hList) {
			if (pnm->code == LVN_COLUMNCLICK) {
				NMLISTVIEW* plv = (NMLISTVIEW*)lParam;
				int col = plv->iSubItem;
				if (g_sortColumn == col) g_sortAscending = !g_sortAscending;
				else { g_sortColumn = col; g_sortAscending = true; }

				ListView_SortItems(g_hList, [](LPARAM l1, LPARAM l2, LPARAM lSort)->int {
					int col = LOWORD(lSort);
					bool asc = HIWORD(lSort) != 0;
					std::wstring* a = reinterpret_cast<std::wstring*>(l1);
					std::wstring* b = reinterpret_cast<std::wstring*>(l2);
					if (!a || !b) return 0;
					std::wstring va;
					std::wstring vb;
					if (col == 0) {
						auto extract_name_no_ext = [](const std::wstring& full)->std::wstring {
							size_t p = full.find_last_of(L"\\/");
							std::wstring name = (p == std::wstring::npos) ? full : full.substr(p + 1);
							size_t d = name.find_last_of(L'.');
							return (d == std::wstring::npos) ? name : name.substr(0, d);
							};
						va = extract_name_no_ext(*a);
						vb = extract_name_no_ext(*b);
					}
					else if (col == 1) {
						size_t pa = a->find_last_of(L"\\/");
						std::wstring na = (pa == std::wstring::npos) ? *a : a->substr(pa + 1);
						size_t da = na.find_last_of(L'.');
						va = (da == std::wstring::npos) ? L"" : na.substr(da);
						size_t pb = b->find_last_of(L"\\/");
						std::wstring nb = (pb == std::wstring::npos) ? *b : b->substr(pb + 1);
						size_t db = nb.find_last_of(L'.');
						vb = (db == std::wstring::npos) ? L"" : nb.substr(db);
					}
					else if (col == 2) {
						ULONGLONG la = GetMediaDurationMs(*a);
						ULONGLONG lb = GetMediaDurationMs(*b);
						if (la < lb) return asc ? -1 : 1;
						if (la > lb) return asc ? 1 : -1;
					}
					else {
						size_t pa = a->find_last_of(L"\\/");
						va = (pa == std::wstring::npos) ? L"" : a->substr(0, pa);
						size_t pb = b->find_last_of(L"\\/");
						vb = (pb == std::wstring::npos) ? L"" : b->substr(0, pb);
					}
					int cmp = _wcsicmp(va.c_str(), vb.c_str());
					return asc ? cmp : -cmp;
					}, MAKELPARAM(g_sortColumn, g_sortAscending ? 1 : 0));
			}
		}

		break;
	}

	default:
		return DefWindowProcW(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow) {
	g_hInst = hInstance;
	OleInitialize(NULL);

	// choose sizes using system metrics for DPI-awareness
	int cxIcon = GetSystemMetrics(SM_CXICON);   // typically 32
	int cyIcon = GetSystemMetrics(SM_CYICON);
	int cxSm = GetSystemMetrics(SM_CXSMICON); // typically 16
	int cySm = GetSystemMetrics(SM_CYSMICON);

	// load icons from resources (IDC_MYICON)
	g_hIconBig = (HICON)LoadImageW(hInstance, MAKEINTRESOURCE(IDI_PLAYLISTDND),
		IMAGE_ICON, cxIcon, cyIcon, LR_DEFAULTCOLOR);
	g_hIconSmall = (HICON)LoadImageW(hInstance, MAKEINTRESOURCE(IDI_SMALL),
		IMAGE_ICON, cxSm, cySm, LR_DEFAULTCOLOR);

	const wchar_t CLASS_NAME[] = L"PlaylistDnDWndClass";
	WNDCLASSEXW wcx = {};
	wcx.cbSize = sizeof(wcx);
	wcx.style = CS_HREDRAW | CS_VREDRAW;
	wcx.lpfnWndProc = WndProc;
	wcx.hInstance = hInstance;
	wcx.lpszClassName = CLASS_NAME;
	wcx.hCursor = LoadCursor(NULL, IDC_ARROW);
	wcx.hIcon = g_hIconBig;
	wcx.hIconSm = g_hIconSmall;
	RegisterClassExW(&wcx);

	HWND hWnd = CreateWindowExW(0, CLASS_NAME, L"Playlist Drag and Drop",
		WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 600, 450,
		NULL, NULL, hInstance, NULL);

	if (!hWnd) return 0;

	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);

	SendMessageW(hWnd, WM_SETICON, ICON_BIG, (LPARAM)g_hIconBig);
	SendMessageW(hWnd, WM_SETICON, ICON_SMALL, (LPARAM)g_hIconSmall);

	MSG msg;
	while (GetMessageW(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessageW(&msg);
	}

	if (g_hIconBig) { DestroyIcon(g_hIconBig);  g_hIconBig = NULL; }
	if (g_hIconSmall) { DestroyIcon(g_hIconSmall); g_hIconSmall = NULL; }

	OleUninitialize();
	return 0;
}
