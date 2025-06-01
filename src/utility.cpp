#include "utility.h" 



std::wstring getAbsolutePath(const std::wstring& relativePath) {
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(nullptr, exePath, MAX_PATH);

    std::filesystem::path exeDir = std::filesystem::path(exePath).parent_path();
    std::filesystem::path absPath = std::filesystem::canonical(exeDir / relativePath);

    return absPath.wstring();
}