#ifndef LOAD_CONFIG_HPP
#define LOAD_CONFIG_HPP
#include <windows.h>
#include <commdlg.h>
#include <iostream>
#include <string>
#include <map>
#include "helper.hpp"
using namespace std;

string select_file()
{
    char filename[MAX_PATH] = "";
    OPENFILENAMEA ofn;
    ZeroMemory(&ofn, sizeof(ofn));

    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = NULL;
    ofn.lpstrFilter = "All Files\0*.*\0Text Files\0*.txt\0";
    ofn.lpstrFile = filename;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    ofn.lpstrTitle = "Select an attachment";

    if (GetOpenFileNameA(&ofn))
    {
        return string(filename);
    }
    else
    {
        return "";
    }
}

bool load_config(string config_file_path, map<string, string> *config)
{
    try
    {
        if (!file_exist(config_file_path.c_str()))
        {
            printf("CONFIG FILE DOES NOT EXIST");
            return false;
        }

        str_vec config_lines = split(read_text_file(config_file_path.c_str()), "\n");
        config_lines.erase(config_lines.begin());
        for (string s : config_lines)
        {
            str_vec pair = split(s, "=");
            (*config)[pair[0]] = pair[1];
        }
    }
    catch (const std::exception &ex)
    {
        std::cerr << "Standard exception: " << ex.what() << std::endl;
        return false;
    }
    catch (...)
    {
        std::cerr << "Unknown exception occurred." << std::endl;
        return false;
    }
    return true;
}

#endif