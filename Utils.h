#pragma once

#include <string>
#include <vector>
#include <shlwapi.h>
#pragma comment(lib, "shlwapi.lib")

inline std::wstring Trim(const std::wstring& s) {
    size_t a = s.find_first_not_of(L" \t\r\n");
    if (a == std::wstring::npos) return L"";
    size_t b = s.find_last_not_of(L" \t\r\n");
    return s.substr(a, b - a + 1);
}

inline std::wstring CombinePath(const std::wstring& baseDir, const std::wstring& relative) {
    if (relative.size() >= 2 && relative[1] == L':') return relative; // absolute
    if (!baseDir.empty() && (relative.size() >= 2 && relative[0] == L'\\' && relative[1] == L'\\')) return relative;
    std::wstring out = baseDir;
    if (!out.empty() && out.back() != L'\\') out.push_back(L'\\');
    out += relative;
    // Normalize: remove "./" and ".." minimally (not full canonicalization)
    // For simplicity, return as-is; Shell consumers will accept relative-ish paths if valid.
    return out;
}

inline std::wstring ResolvePath(const std::wstring& baseDir, const std::wstring& entry)
{
    std::wstring candidate = entry;
    // If relative, combine
    if (!(entry.size() >= 2 && entry[1] == L':') && !(entry.size() >= 2 && entry[0] == L'\\' && entry[1] == L'\\')) {
        std::wstring combined = baseDir;
        if (!combined.empty() && combined.back() != L'\\') combined.push_back(L'\\');
        combined += entry;
        candidate = combined;
    }
    // Get full path
    DWORD needed = GetFullPathNameW(candidate.c_str(), 0, NULL, NULL);
    if (needed) {
        std::wstring full;
        full.resize(needed);
        GetFullPathNameW(candidate.c_str(), needed, &full[0], NULL);
        // Remove trailing null added by GetFullPathNameW
        if (!full.empty() && full.back() == L'\0') full.pop_back();
        // Canonicalize (removes .., .)
        wchar_t buf[MAX_PATH];
        if (PathCanonicalizeW(buf, full.c_str())) return std::wstring(buf);
        return full;
    }
    return candidate;
}

// Append a single path to the vector if it is a file (optionally check existence)
inline void AddFileIfExists(std::vector<std::wstring>& out, const std::wstring& path) {
	if (GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES) {
		out.push_back(path);
	}
}
