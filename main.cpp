#include <iostream>
#include "load_config.hpp"
#include "COM/COM.hpp"
int main()
{
    map<string, string> config;
    load_config("C:\\Users\\sfcuser\\Desktop\\NEIS_TOOLS\\NEIC_TOOLS\\config.config", &config);
    CoInitialize(NULL);
    // outlook_send("fangzhou.ye@fii-na.com", "", "TEST SUBJECT", "TEST BODY", select_file());

    IDispatch *Excelp = get_excel_app();
    IDispatch *ExcelWorkbooksp = get_workbooks(Excelp);
    IDispatch *ExcelWorkbookp = get_workbook(ExcelWorkbooksp, select_file());
    IDispatch *Excelsheetp = get_sheet(ExcelWorkbookp, 1);
    CoUninitialize();
    return 0;
}