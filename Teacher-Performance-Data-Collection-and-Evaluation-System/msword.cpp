#include "pch.h"
#include "msword.h"

CMSWord::CMSWord()
    : m_pWordApp(NULL)
    , m_pDocuments(NULL)
    , m_pDocument(NULL)
    , m_pSelection(NULL)
    , m_pTable(NULL)
    , m_bInit(false)
{
}

CMSWord::~CMSWord()
{
    CloseWord();
}

bool CMSWord::CreateWordInstance()
{
    if (m_bInit) return true;

    // 初始化COM
    if (FAILED(::CoInitialize(NULL)))
        return false;

    // 创建Word应用程序
    CLSID clsid;
    if (FAILED(CLSIDFromProgID(L"Word.Application", &clsid)))
        return false;

    if (FAILED(::CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&m_pWordApp)))
        return false;

    m_bInit = true;
    return true;
}

bool CMSWord::CreateNewDocument()
{
    if (!m_bInit || !m_pWordApp)
        return false;

    // 获取Documents集合
    DISPID dispid;
    OLECHAR* name = L"Documents";
    if (FAILED(m_pWordApp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    if (FAILED(m_pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL)))
        return false;

    m_pDocuments = result.pdispVal;

    // 添加新文档
    name = L"Add";
    if (FAILED(m_pDocuments->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    if (FAILED(m_pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL)))
        return false;

    m_pDocument = result.pdispVal;

    // 获取Selection对象
    name = L"Selection";
    if (FAILED(m_pWordApp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    if (FAILED(m_pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL)))
        return false;

    m_pSelection = result.pdispVal;
    if (!m_pSelection) {
        AfxMessageBox(_T("m_pSelection 初始化失败"));
        return false;
    }
    return true;
}

bool CMSWord::AddTitle(const CString& title)
{
    if (!m_pSelection)
        return false;

    // 设置文本
    DISPID dispid;
    OLECHAR* name = L"TypeText";
    if (FAILED(m_pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    VARIANT args[1];
    args[0].vt = VT_BSTR;
    args[0].bstrVal = title.AllocSysString();

    DISPPARAMS dp = { args, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);
    SysFreeString(args[0].bstrVal);

    // 添加段落标记
    name = L"TypeParagraph";
    if (FAILED(m_pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    dp.cArgs = 0;
    return SUCCEEDED(m_pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL));
}

bool CMSWord::AddParagraph(const CString& text)
{
    if (!m_pSelection)
        return false;

    // 设置文本
    DISPID dispid;
    OLECHAR* name = L"TypeText";
    if (FAILED(m_pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    VARIANT args[1];
    args[0].vt = VT_BSTR;
    args[0].bstrVal = text.AllocSysString();

    DISPPARAMS dp = { args, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);
    SysFreeString(args[0].bstrVal);

    // 添加段落标记
    name = L"TypeParagraph";
    if (FAILED(m_pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    dp.cArgs = 0;
    return SUCCEEDED(m_pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL));
}

bool CMSWord::CreateTable(int rows, int cols)
{
    if (!m_pSelection) {
        AfxMessageBox(_T("m_pSelection 未初始化"));
        return false;
    }

    if (rows <= 0 || cols <= 0) {
        AfxMessageBox(_T("行数或列数无效"));
        return false;
    }

    // 获取 Tables 集合
    DISPID dispid;
    OLECHAR* name = L"Tables";
    if (FAILED(m_pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid))) {
        AfxMessageBox(_T("获取 Tables 属性 ID 失败"));
        return false;
    }

    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    if (FAILED(m_pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL))) {
        AfxMessageBox(_T("调用 Tables 属性失败"));
        return false;
    }

    LPDISPATCH pTables = result.pdispVal;
    if (!pTables) {
        AfxMessageBox(_T("获取 Tables 集合失败"));
        return false;
    }

    // 获取 Selection 的 Range 作为插入位置
    VARIANT rangeVar;
    VariantInit(&rangeVar);
    name = L"Range";
    if (FAILED(m_pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid))) {
        AfxMessageBox(_T("获取 Range 属性 ID 失败"));
        pTables->Release();
        return false;
    }

    if (FAILED(m_pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &rangeVar, NULL, NULL))) {
        AfxMessageBox(_T("调用 Range 属性失败"));
        pTables->Release();
        return false;
    }

    // 调用 Add 方法插入表格
    name = L"Add";
    if (FAILED(pTables->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid))) {
        AfxMessageBox(_T("获取 Add 方法 ID 失败"));
        pTables->Release();
        return false;
    }

    VARIANT args[3];
    args[0].vt = VT_I4;           // 列数
    args[0].lVal = cols;
    args[1].vt = VT_I4;           // 行数
    args[1].lVal = rows;
    args[2] = rangeVar;           // 插入位置

    dp.rgvarg = args;
    dp.cArgs = 3;

    HRESULT hr = pTables->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);
    if (FAILED(hr)) {
        CString errorMessage;
        errorMessage.Format(_T("调用 Add 方法失败，错误代码: 0x%08X"), hr);
        AfxMessageBox(errorMessage);
        VariantClear(&rangeVar);
        pTables->Release();
        return false;
    }

    m_pTable = result.pdispVal;
    VariantClear(&rangeVar);
    pTables->Release();

    if (!m_pTable) {
        AfxMessageBox(_T("创建表格失败，m_pTable 为 NULL"));
        return false;
    }

    return true;
}

bool CMSWord::SetTableCell(int row, int col, const CString& text)
{
    if (!m_pTable)
    {
        OutputDebugString(_T("表格未初始化，无法设置单元格内容\n"));
        return false;
    }

    // 调试信息
    CString msg;
    msg.Format(_T("准备设置单元格内容，row = %d, col = %d, text = %s\n"), row, col, text);
    OutputDebugString(msg);

    // 获取 Cell 方法的 ID
    DISPID dispid;
    OLECHAR* name = L"Cell";
    if (FAILED(m_pTable->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
    {
        OutputDebugString(_T("获取 Cell 方法的 ID 失败\n"));
        return false;
    }

    // 调用 Cell 方法获取单元格对象
    VARIANT args[2];
    args[0].vt = VT_I4;
    args[0].lVal = col;
    args[1].vt = VT_I4;
    args[1].lVal = row;

    DISPPARAMS dp = { args, NULL, 2, 0 };
    VARIANT result;
    VariantInit(&result);

    if (FAILED(m_pTable->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL)))
    {
        OutputDebugString(_T("调用 Cell 方法失败，无法获取单元格对象\n"));
        return false;
    }

    // 检查 Cell 对象
    LPDISPATCH pCell = result.pdispVal;
    if (!pCell)
    {
        OutputDebugString(_T("获取单元格失败，pCell 为 NULL\n"));
        return false;
    }

    // 获取 Range 属性
    name = L"Range";
    if (FAILED(pCell->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
    {
        OutputDebugString(_T("获取 Range 属性的 ID 失败\n"));
        pCell->Release();
        return false;
    }

    VariantInit(&result);
    if (FAILED(pCell->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL)))
    {
        OutputDebugString(_T("调用 Range 属性失败\n"));
        pCell->Release();
        return false;
    }

    // 检查 Range 对象
    LPDISPATCH pRange = result.pdispVal;
    if (!pRange)
    {
        OutputDebugString(_T("获取单元格范围失败，pRange 为 NULL\n"));
        pCell->Release();
        return false;
    }

    // 设置单元格内容
    name = L"Text";
    if (FAILED(pRange->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
    {
        OutputDebugString(_T("获取 Text 属性的 ID 失败\n"));
        pRange->Release();
        pCell->Release();
        return false;
    }

    VARIANT textVar;
    textVar.vt = VT_BSTR;
    textVar.bstrVal = text.AllocSysString();

    DISPPARAMS dpText = { &textVar, NULL, 1, 0 };
    HRESULT hr = pRange->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpText, NULL, NULL, NULL);

    // 释放内存
    SysFreeString(textVar.bstrVal);
    pRange->Release();
    pCell->Release();

    if (FAILED(hr))
    {
        msg.Format(_T("设置单元格内容失败，错误代码：0x%08X\n"), hr);
        OutputDebugString(msg);
        return false;
    }

    return true;
}

bool CMSWord::SaveAs(const CString& filePath)
{
    if (!m_pDocument)
        return false;

    DISPID dispid;
    OLECHAR* name = L"SaveAs";
    if (FAILED(m_pDocument->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    VARIANT args[1];
    args[0].vt = VT_BSTR;
    args[0].bstrVal = filePath.AllocSysString();

    DISPPARAMS dp = { args, NULL, 1, 0 };
    HRESULT hr = m_pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, NULL, NULL, NULL);
    SysFreeString(args[0].bstrVal);

    return SUCCEEDED(hr);
}

void CMSWord::CloseWord()
{
    if (m_pWordApp)
    {
        DISPID dispid;
        OLECHAR* name = L"Quit";
        if (SUCCEEDED(m_pWordApp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        {
            DISPPARAMS dp = { NULL, NULL, 0, 0 };
            m_pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, NULL, NULL, NULL);
        }
        m_pWordApp->Release();
        m_pWordApp = NULL;
    }
    ::CoUninitialize();
}
