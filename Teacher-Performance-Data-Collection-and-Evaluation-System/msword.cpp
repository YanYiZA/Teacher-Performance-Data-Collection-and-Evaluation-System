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

    // ��ʼ��COM
    if (FAILED(::CoInitialize(NULL)))
        return false;

    // ����WordӦ�ó���
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

    // ��ȡDocuments����
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

    // ������ĵ�
    name = L"Add";
    if (FAILED(m_pDocuments->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    if (FAILED(m_pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL)))
        return false;

    m_pDocument = result.pdispVal;

    // ��ȡSelection����
    name = L"Selection";
    if (FAILED(m_pWordApp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    if (FAILED(m_pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL)))
        return false;

    m_pSelection = result.pdispVal;
    if (!m_pSelection) {
        AfxMessageBox(_T("m_pSelection ��ʼ��ʧ��"));
        return false;
    }
    return true;
}

bool CMSWord::AddTitle(const CString& title)
{
    if (!m_pSelection)
        return false;

    // �����ı�
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

    // ��Ӷ�����
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

    // �����ı�
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

    // ��Ӷ�����
    name = L"TypeParagraph";
    if (FAILED(m_pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
        return false;

    dp.cArgs = 0;
    return SUCCEEDED(m_pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL));
}

bool CMSWord::CreateTable(int rows, int cols)
{
    if (!m_pSelection) {
        AfxMessageBox(_T("m_pSelection δ��ʼ��"));
        return false;
    }

    if (rows <= 0 || cols <= 0) {
        AfxMessageBox(_T("������������Ч"));
        return false;
    }

    // ��ȡ Tables ����
    DISPID dispid;
    OLECHAR* name = L"Tables";
    if (FAILED(m_pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid))) {
        AfxMessageBox(_T("��ȡ Tables ���� ID ʧ��"));
        return false;
    }

    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    if (FAILED(m_pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL))) {
        AfxMessageBox(_T("���� Tables ����ʧ��"));
        return false;
    }

    LPDISPATCH pTables = result.pdispVal;
    if (!pTables) {
        AfxMessageBox(_T("��ȡ Tables ����ʧ��"));
        return false;
    }

    // ��ȡ Selection �� Range ��Ϊ����λ��
    VARIANT rangeVar;
    VariantInit(&rangeVar);
    name = L"Range";
    if (FAILED(m_pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid))) {
        AfxMessageBox(_T("��ȡ Range ���� ID ʧ��"));
        pTables->Release();
        return false;
    }

    if (FAILED(m_pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &rangeVar, NULL, NULL))) {
        AfxMessageBox(_T("���� Range ����ʧ��"));
        pTables->Release();
        return false;
    }

    // ���� Add ����������
    name = L"Add";
    if (FAILED(pTables->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid))) {
        AfxMessageBox(_T("��ȡ Add ���� ID ʧ��"));
        pTables->Release();
        return false;
    }

    VARIANT args[3];
    args[0].vt = VT_I4;           // ����
    args[0].lVal = cols;
    args[1].vt = VT_I4;           // ����
    args[1].lVal = rows;
    args[2] = rangeVar;           // ����λ��

    dp.rgvarg = args;
    dp.cArgs = 3;

    HRESULT hr = pTables->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);
    if (FAILED(hr)) {
        CString errorMessage;
        errorMessage.Format(_T("���� Add ����ʧ�ܣ��������: 0x%08X"), hr);
        AfxMessageBox(errorMessage);
        VariantClear(&rangeVar);
        pTables->Release();
        return false;
    }

    m_pTable = result.pdispVal;
    VariantClear(&rangeVar);
    pTables->Release();

    if (!m_pTable) {
        AfxMessageBox(_T("�������ʧ�ܣ�m_pTable Ϊ NULL"));
        return false;
    }

    return true;
}

bool CMSWord::SetTableCell(int row, int col, const CString& text)
{
    if (!m_pTable)
    {
        OutputDebugString(_T("���δ��ʼ�����޷����õ�Ԫ������\n"));
        return false;
    }

    // ������Ϣ
    CString msg;
    msg.Format(_T("׼�����õ�Ԫ�����ݣ�row = %d, col = %d, text = %s\n"), row, col, text);
    OutputDebugString(msg);

    // ��ȡ Cell ������ ID
    DISPID dispid;
    OLECHAR* name = L"Cell";
    if (FAILED(m_pTable->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
    {
        OutputDebugString(_T("��ȡ Cell ������ ID ʧ��\n"));
        return false;
    }

    // ���� Cell ������ȡ��Ԫ�����
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
        OutputDebugString(_T("���� Cell ����ʧ�ܣ��޷���ȡ��Ԫ�����\n"));
        return false;
    }

    // ��� Cell ����
    LPDISPATCH pCell = result.pdispVal;
    if (!pCell)
    {
        OutputDebugString(_T("��ȡ��Ԫ��ʧ�ܣ�pCell Ϊ NULL\n"));
        return false;
    }

    // ��ȡ Range ����
    name = L"Range";
    if (FAILED(pCell->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
    {
        OutputDebugString(_T("��ȡ Range ���Ե� ID ʧ��\n"));
        pCell->Release();
        return false;
    }

    VariantInit(&result);
    if (FAILED(pCell->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL)))
    {
        OutputDebugString(_T("���� Range ����ʧ��\n"));
        pCell->Release();
        return false;
    }

    // ��� Range ����
    LPDISPATCH pRange = result.pdispVal;
    if (!pRange)
    {
        OutputDebugString(_T("��ȡ��Ԫ��Χʧ�ܣ�pRange Ϊ NULL\n"));
        pCell->Release();
        return false;
    }

    // ���õ�Ԫ������
    name = L"Text";
    if (FAILED(pRange->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid)))
    {
        OutputDebugString(_T("��ȡ Text ���Ե� ID ʧ��\n"));
        pRange->Release();
        pCell->Release();
        return false;
    }

    VARIANT textVar;
    textVar.vt = VT_BSTR;
    textVar.bstrVal = text.AllocSysString();

    DISPPARAMS dpText = { &textVar, NULL, 1, 0 };
    HRESULT hr = pRange->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpText, NULL, NULL, NULL);

    // �ͷ��ڴ�
    SysFreeString(textVar.bstrVal);
    pRange->Release();
    pCell->Release();

    if (FAILED(hr))
    {
        msg.Format(_T("���õ�Ԫ������ʧ�ܣ�������룺0x%08X\n"), hr);
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
