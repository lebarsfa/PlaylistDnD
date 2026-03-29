#pragma once

#include <fstream>
#include <sstream>
#include <algorithm>
#include "Utils.h"

// Percent-encode a UTF-16 URL string (encode space, control chars, and reserved/unsafe chars).
inline std::wstring PercentEncodeUrlW(const std::wstring& src)
{
    auto is_unreserved = [](wchar_t c)->bool {
        if ((c >= L'a' && c <= L'z') || (c >= L'A' && c <= L'Z') || (c >= L'0' && c <= L'9')) return true;
        switch (c) {
        case L'-': case L'.': case L'_': case L'~': return true;
        default: return false;
        }
    };

    std::wstring out;
    out.reserve(src.size() * 3);
    for (wchar_t wc : src) {
        // encode control, space, non-ASCII, and reserved characters
        if (is_unreserved(wc)) {
            out.push_back(wc);
        }
        else if (wc < 0x80) {
            // ASCII -> percent-encode single byte
            wchar_t buf[4];
            swprintf_s(buf, L"%%%02X", (unsigned int)wc & 0xFF);
            out += buf;
        }
        else {
            // Non-ASCII: convert wchar to UTF-8 bytes then percent-encode each byte
            char utf8[8] = {};
            int n = WideCharToMultiByte(CP_UTF8, 0, &wc, 1, utf8, (int)sizeof(utf8), NULL, NULL);
            for (int i = 0; i < n; ++i) {
                wchar_t buf[4];
                swprintf_s(buf, L"%%%02X", (unsigned char)utf8[i]);
                out += buf;
            }
        }
    }
    return out;
}

// Convert UTF-16 string to ANSI (current code page) in a std::string
inline std::string WideToAnsi(const std::wstring& w)
{
    int needed = WideCharToMultiByte(CP_ACP, 0, w.c_str(), (int)w.size(), NULL, 0, NULL, NULL);
    if (needed <= 0) return std::string();
    std::string out;
    out.resize(needed);
    WideCharToMultiByte(CP_ACP, 0, w.c_str(), (int)w.size(), &out[0], needed, NULL, NULL);
    return out;
}

// helper: percent-decode bytes into UTF-8 then to wstring
inline std::wstring PercentDecodeUtf8ToW(const std::wstring& s)
{
	// If no percent, return original
	if (s.find(L'%') == std::wstring::npos) return s;

	// Build a byte buffer from percent-decoding; for non-% characters, encode as UTF-8 bytes
	std::string bytes;
	bytes.reserve(s.size());
	for (size_t i = 0; i < s.size(); ++i) {
		wchar_t wc = s[i];
		if (wc == L'%') {
			// expect two hex digits (in narrow chars)
			if (i + 2 < s.size()) {
				wchar_t h1 = s[i+1], h2 = s[i+2];
				auto hexVal = [](wchar_t c)->int {
					if (c >= L'0' && c <= L'9') return c - L'0';
					if (c >= L'A' && c <= L'F') return 10 + (c - L'A');
					if (c >= L'a' && c <= L'f') return 10 + (c - L'a');
					return -1;
					};
				int v1 = hexVal(h1), v2 = hexVal(h2);
				if (v1 >= 0 && v2 >= 0) {
					unsigned char b = (unsigned char)((v1 << 4) | v2);
					bytes.push_back((char)b);
					i += 2;
					continue;
				}
			}
			// malformed percent, treat '%' literally
			bytes.push_back('%');
		}
		else {
			// convert wchar_t to UTF-8 bytes
			char buf[8] = {};
			int n = WideCharToMultiByte(CP_UTF8, 0, &wc, 1, buf, (int)sizeof(buf), NULL, NULL);
			if (n > 0) bytes.append(buf, buf + n);
		}
	}

	// convert UTF-8 bytes to wstring
	if (bytes.empty()) return std::wstring();
	int needed = MultiByteToWideChar(CP_UTF8, 0, bytes.data(), (int)bytes.size(), NULL, 0);
	if (needed <= 0) return std::wstring();
	std::wstring out;
	out.resize(needed);
	MultiByteToWideChar(CP_UTF8, 0, bytes.data(), (int)bytes.size(), &out[0], needed);
	return out;
}

// helper: normalize an M3U entry line into a Windows path or URL
inline std::wstring NormalizeM3UEntry(const std::wstring& raw)
{
	if (raw.empty()) return raw;

	// Trim whitespace already done by caller; work on a copy
	std::wstring s = raw;

	// If it's a file URI (file://), strip scheme and host if present
	const std::wstring fileScheme = L"file://";
	if (_wcsnicmp(s.c_str(), fileScheme.c_str(), fileScheme.size()) == 0) {
		// file://localhost/C:/path or file:///C:/path or file://server/share/path
		std::wstring tail = s.substr(fileScheme.size());
		// If it starts with '/', remove a single leading slash for local paths like file:///C:/...
		if (!tail.empty() && tail[0] == L'/') {
			// if next is drive letter, remove leading slash
			if (tail.size() >= 3 && iswalpha(tail[1]) && tail[2] == L':') {
				tail = tail.substr(1);
			}
		}
		// convert forward slashes to backslashes
		for (auto& ch : tail) if (ch == L'/') ch = L'\\';
		// percent-decode any %XX sequences
		tail = PercentDecodeUtf8ToW(tail);
		return tail;
	}

	// If it looks like an HTTP/HTTPS URL, keep as-is (or optionally skip)
	const std::wstring http = L"http://";
	const std::wstring https = L"https://";
	if (_wcsnicmp(s.c_str(), http.c_str(), http.size()) == 0 ||
		_wcsnicmp(s.c_str(), https.c_str(), https.size()) == 0) {
		// leave URL unchanged (you may choose to store it as-is)
		return s;
	}

	// For plain paths: percent-decode then normalize slashes
	std::wstring decoded = PercentDecodeUtf8ToW(s);
	for (auto& ch : decoded) if (ch == L'/') ch = L'\\';
	return decoded;
}

// Read file bytes, detect BOM, convert to UTF-16 wstring lines
inline std::vector<std::wstring> ReadM3ULines(const std::wstring& path)
{
	std::vector<std::wstring> out;
	std::ifstream in(path, std::ios::binary);
	if (!in) return out;

	// Read all bytes
	std::string bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());

	// Detect BOM
	enum Enc { ENC_UTF8, ENC_UTF8_BOM, ENC_UTF16_LE, ENC_UTF16_BE, ENC_ANSI };
	Enc enc = ENC_ANSI;
	if (bytes.size() >= 3 &&
		(unsigned char)bytes[0] == 0xEF &&
		(unsigned char)bytes[1] == 0xBB &&
		(unsigned char)bytes[2] == 0xBF) {
		enc = ENC_UTF8_BOM;
	}
	else if (bytes.size() >= 2 &&
		(unsigned char)bytes[0] == 0xFF &&
		(unsigned char)bytes[1] == 0xFE) {
		enc = ENC_UTF16_LE;
	}
	else if (bytes.size() >= 2 &&
		(unsigned char)bytes[0] == 0xFE &&
		(unsigned char)bytes[1] == 0xFF) {
		enc = ENC_UTF16_BE;
	}
	else {
		// Heuristic: if bytes contain a 0x00 every other byte, treat as UTF-16 LE
		bool maybeUtf16 = false;
		for (size_t i = 0; i + 1 < bytes.size() && i < 200; i += 2) {
			if (bytes[i+1] == 0x00) { maybeUtf16 = true; break; }
		}
		enc = maybeUtf16 ? ENC_UTF16_LE : ENC_UTF8;
	}

	std::wstring wide;
	if (enc == ENC_UTF16_LE) {
		// Skip BOM if present
		size_t offset = (bytes.size() >= 2 && (unsigned char)bytes[0] == 0xFF && (unsigned char)bytes[1] == 0xFE) ? 2 : 0;
		size_t wcharCount = (bytes.size() - offset) / 2;
		wide.resize(wcharCount);
		memcpy(&wide[0], bytes.data() + offset, wcharCount * sizeof(wchar_t));
	}
	else if (enc == ENC_UTF16_BE) {
		// Convert BE to LE
		size_t offset = (bytes.size() >= 2 && (unsigned char)bytes[0] == 0xFE && (unsigned char)bytes[1] == 0xFF) ? 2 : 0;
		size_t wc = (bytes.size() - offset) / 2;
		wide.resize(wc);
		for (size_t i = 0; i < wc; ++i) {
			unsigned char hi = bytes[offset + i*2];
			unsigned char lo = bytes[offset + i*2 + 1];
			wide[i] = (wchar_t)((lo) | (hi << 8));
		}
	}
	else {
		// UTF-8 (with or without BOM) or ANSI: convert to UTF-16 using MultiByteToWideChar
		size_t offset = 0;
		if (enc == ENC_UTF8_BOM) offset = 3;
		int needed = MultiByteToWideChar(CP_UTF8, 0, bytes.data() + offset, (int)(bytes.size() - offset), NULL, 0);
		if (needed > 0) {
			wide.resize(needed);
			MultiByteToWideChar(CP_UTF8, 0, bytes.data() + offset, (int)(bytes.size() - offset), &wide[0], needed);
		}
		else {
			// Fallback: treat as ANSI
			int n2 = MultiByteToWideChar(CP_ACP, 0, bytes.data(), (int)bytes.size(), NULL, 0);
			if (n2 > 0) {
				wide.resize(n2);
				MultiByteToWideChar(CP_ACP, 0, bytes.data(), (int)bytes.size(), &wide[0], n2);
			}
		}
	}

	// Split into lines (handle CRLF, LF)
	size_t start = 0;
	for (size_t i = 0; i <= wide.size(); ++i) {
		if (i == wide.size() || wide[i] == L'\n') {
			size_t len = (i > 0 && wide[i-1] == L'\r') ? (i - 1 - start) : (i - start);
			std::wstring line = wide.substr(start, len);

			// Trim whitespace
			size_t a = line.find_first_not_of(L" \t\r\n");
			if (a != std::wstring::npos) {
				size_t b = line.find_last_not_of(L" \t\r\n");
				line = line.substr(a, b - a + 1);
			}
			else line.clear();

			// Skip empty lines and comment/tag lines that start with '#'
			if (!line.empty()) {
				if (line[0] == L'#') {
					// skip tags/comments like #EXTM3U, #EXTINF, #EXT-X-...
					// advance to next line without adding to output
				}
				else {
					// Normalize percent-encoding and slashes
					std::wstring normalized = NormalizeM3UEntry(line);
					if (!normalized.empty()) out.push_back(normalized);
				}
			}

			start = i + 1;
		}
	}
	return out;
}
