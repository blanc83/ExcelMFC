#include "pch.h"
#include "framework.h"
#include "ExcelMFC.h"
#include "ExcelMFCDlg.h"
#include "afxdialogex.h"
#include <fstream>
#include <sstream>
#include <algorithm>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CExcelMFCDialog::CExcelMFCDialog(CWnd* pParent /*=nullptr*/)
    : CDialogEx(IDD_EXCELMFC_DIALOG, pParent)
{
}

CExcelMFCDialog::~CExcelMFCDialog()
{
    // 필요한 정리 작업
}

void CExcelMFCDialog::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Text(pDX, IDC_EDIT1, SName);
    DDX_Control(pDX, IDC_COMBO1, m_combobox1);
    DDX_Control(pDX, IDC_EDIT_VALUE4, m_editvalue4);
    DDX_Control(pDX, IDC_EDIT_VALUE5, m_editvalue5);
    DDX_Control(pDX, IDC_EDIT_VALUE6, m_editvalue6);
    DDX_Control(pDX, IDC_EDIT_VALUE8, m_editvalue8);
    DDX_Control(pDX, IDC_EDIT_VALUE7, m_editvalue7);
    DDX_Control(pDX, IDC_EDIT3, apply1);
    DDX_Control(pDX, IDC_EDIT4, apply2);
    DDX_Control(pDX, IDC_EDIT6, apply3);
    DDX_Control(pDX, IDC_EDIT7, apply4);
    DDX_Control(pDX, IDC_EDIT2, apply5);
    DDX_Control(pDX, IDC_EDIT8, apply6);
}
BEGIN_MESSAGE_MAP(CExcelMFCDialog, CDialogEx)
    ON_BN_CLICKED(IDC_BUTTON1, &CExcelMFCDialog::OnBnClickedAddName)  // 이름 추가 버튼
    ON_BN_CLICKED(IDC_OnBnClickedBrowse, &CExcelMFCDialog::OnBnClickedBrowse)  // 파일 열기 버튼
    ON_BN_CLICKED(IDC_BUTTON3, &CExcelMFCDialog::OnBnClickButton3)
    ON_BN_CLICKED(IDC_BUTTON_BACKUP, &CExcelMFCDialog::OnBnClickedBackup)
    ON_BN_CLICKED(IDC_BUTTON4, &CExcelMFCDialog::OnBnClickedApply)
    ON_BN_CLICKED(IDC_BUTTON_SORT, &CExcelMFCDialog::OnBnClickedButtonSort)
END_MESSAGE_MAP()

// CSV 파일 읽기
void CExcelMFCDialog::ReadCSVFile(const CString& _FilePath, std::vector<std::vector<CString>>& data)
{
    std::ifstream file(CW2A(_FilePath.GetString()), std::ios::in);

    if (!file.is_open()) {
        AfxMessageBox(_T("파일을 열 수 없습니다."));
        return;
    }
    data.clear();
    std::string line;
    while (std::getline(file, line)) {
        std::vector<CString> row;
        std::stringstream ss(line);
        std::string cell;

        while (std::getline(ss, cell, ',')) {
            row.push_back(CString(cell.c_str()));
        }

        while (row.size() < 1) {
            row.push_back(_T(""));
        }

        data.push_back(row);  // 행 추가
    }

    file.close();
}

void CExcelMFCDialog::BackupCSVFile(const CString& OriginFile)
{
    CTime __CURTIME = CTime::GetCurrentTime();
    CString dateStr = __CURTIME.Format(_T("%Y-%m-%d_%H-%M-%S"));  // 형식: yyyy-mm-dd_hh-mm-ss

    CString BackUpFileName;
    BackUpFileName.Format(_T("%s_backup_%s.csv"), OriginFile.Left(OriginFile.GetLength() - 4), dateStr); // ".csv" 확장자를 제외하고 날짜 추가

    if (_taccess(OriginFile, 0) == -1) {
        AfxMessageBox(_T("원본 파일이 존재하지 않습니다."));
        return;
    }

    if (CopyFile(OriginFile, BackUpFileName, FALSE)) {
        CString successMessage;
        successMessage.Format(_T("백업 파일이 생성되었습니다: %s"), BackUpFileName);
        AfxMessageBox(successMessage);
    }
    else {
        AfxMessageBox(_T("백업 파일 생성 실패."));
    }
}

void CExcelMFCDialog::OnBnClickedBackup()
{
    if (FilePath.IsEmpty()) {
        AfxMessageBox(_T("CSV 파일을 먼저 열어주세요."));
        return;
    }

    // 파일 백업 처리
    BackupCSVFile(FilePath);
}


void CExcelMFCDialog::WriteCSVFile(const CString& __FilePath, const std::vector<std::vector<CString>>& data)
{
    std::ofstream file(static_cast<const char*>(CStringA(__FilePath)));

    if (!file.is_open()) {
        AfxMessageBox(_T("CSV 파일을 저장할 수 없습니다."));
        return;
    }

    for (const auto& row : data) {
        for (size_t i = 0; i < row.size(); ++i) {
            file << CStringA(row[i]);
            if (i < row.size() - 1)
                file << ",";
        }
        file << "\n";
    }

    file.close();
}

void CExcelMFCDialog::OnBnClickedAddName()
{
    if (FilePath.IsEmpty()) {
        AfxMessageBox(_T("csv 파일을 먼저 열어주세요."));
        return;
    }

    UpdateData(TRUE);

    if (SName.IsEmpty()) {
        AfxMessageBox(_T("이름을 입력하세요"));
        return;
    }

    std::vector<std::vector<CString>> data;
    ReadCSVFile(FilePath, data);

    // 중복 이름 확인
    bool SameN = false;
    for (const auto& row : data) {
        if (!row.empty() && row[0] == SName) {
            SameN = true;
            break;
        }
    }

    if (SameN) {
        AfxMessageBox(_T("중복된 이름이 있습니다."));
        return;
    }

    bool AddName = false;
    for (int i = 1; i < 23; i++) {
        if (i >= data.size()) {
            std::vector<CString> newRow = { SName, _T("0"), _T("0"), _T("0"), _T("0"), _T("0") }; // 초기값을 0으로 설정
            data.push_back(newRow);
            AddName = true;
            break;
        }
        else if (data[i].size() < 1 || data[i][0].IsEmpty()) {
            data[i][0] = SName;
            data[i].resize(6, _T("0")); // B2, C2, D2, E2, F2 값을 0으로 설정
            AddName = true;
            break;
        }
    }

    if (!AddName) {
        AfxMessageBox(_T("이름을 추가할 공간이 없습니다."));
    }
    else {
        WriteCSVFile(FilePath, data);

        m_combobox1.ResetContent();

        // A1을 제외한 이름들만 ComboBox에 추가
        bool A1R = true;
        for (const auto& row : data) {
            if (A1R) {
                A1R = false;
                continue;
            }

            if (!row.empty() && !row[0].IsEmpty()) {
                m_combobox1.AddString(row[0]);
            }
        }
        AfxMessageBox(_T("이름이 추가되었습니다."));
    }
}

// 파일 열기
void CExcelMFCDialog::OnBnClickedBrowse()
{
    CFileDialog dlg(TRUE, _T("csv"), NULL, OFN_FILEMUSTEXIST, _T("CSV Files (*.csv)|*.csv|"));
    if (dlg.DoModal() == IDOK) {
        FilePath = dlg.GetPathName();

        CStatic* _StaticText = reinterpret_cast<CStatic*>(GetDlgItem(IDC_STATIC_PATH));
        if (_StaticText) {
            _StaticText->SetWindowText(FilePath);
        }

        AfxMessageBox(_T("파일 적용 :\n") + FilePath);

        std::vector<std::vector<CString>> data;
        ReadCSVFile(FilePath, data);

        m_combobox1.ResetContent();

        // 첫 번째 행 건너뛰기
        bool FROW = true;
        for (const auto& row : data) {
            if (FROW) {
                FROW = false;
                continue;
            }

            if (!row.empty() && !row[0].IsEmpty()) {
                m_combobox1.AddString(row[0]); // 첫 번째 열 데이터를 콤보박스에 추가
            }
        }
    }
}


void CExcelMFCDialog::OnBnClickButton3()
{
    // 읽기 전용 설정
    m_editvalue4.SetReadOnly(TRUE);
    m_editvalue5.SetReadOnly(TRUE);
    m_editvalue6.SetReadOnly(TRUE);
    m_editvalue8.SetReadOnly(TRUE);
    m_editvalue7.SetReadOnly(TRUE);

    int index = m_combobox1.GetCurSel();
    if (index == CB_ERR) {
        AfxMessageBox(_T("이름을 선택해주세요."));
        return;
    }

    CString SelName;
    m_combobox1.GetLBText(index, SelName);

    std::vector<std::vector<CString>> data;
    ReadCSVFile(FilePath, data);
    bool NameFound = false;
    int totalSum = 0;

    for (auto& row : data) {
        if (!row.empty() && row[0] == SelName) {
            // 기존 값 불러오기
            CString value4 = (row.size() > 1) ? row[1] : _T("0");
            CString value5 = (row.size() > 2) ? row[2] : _T("0");
            CString value6 = (row.size() > 3) ? row[3] : _T("0");
            CString value8 = (row.size() > 4) ? row[4] : _T("0");
            CString existingSum = (row.size() > 5) ? row[5] : _T("0");

            // 텍스트 박스에 기존 값 표시
            m_editvalue4.SetWindowText(value4);  // B2
            m_editvalue5.SetWindowText(value5);  // C2
            m_editvalue6.SetWindowText(value6);  // D2
            m_editvalue8.SetWindowText(value8);  // E2
            m_editvalue7.SetWindowText(existingSum);  // F2(합계)

            NameFound = true;
            break;
        }
    }

    if (!NameFound) {
        AfxMessageBox(_T("Not Found Data."));
    }
    else {
        WriteCSVFile(FilePath, data);
    }
}


void CExcelMFCDialog::OnBnClickedApply()
{
    int index = m_combobox1.GetCurSel();
    if (index == CB_ERR) {
        AfxMessageBox(_T("이름을 선택해주세요."));
        return;
    }

    CString _SelName;
    m_combobox1.GetLBText(index, _SelName);

    std::vector<std::vector<CString>> data;
    ReadCSVFile(FilePath, data);

    bool NameFound = false;
    for (auto& row : data) {
        if (!row.empty() && row[0] == _SelName) {
            // 기존 값 읽기
            int OldVal1 = (row.size() > 1) ? _ttoi(row[1]) : 0;
            int OldVal2 = (row.size() > 2) ? _ttoi(row[2]) : 0;
            int OldVal3 = (row.size() > 3) ? _ttoi(row[3]) : 0;
            int OldVal4 = (row.size() > 4) ? _ttoi(row[4]) : 0;
            int OldF1 = (row.size() > 5) ? _ttoi(row[5]) : 0;

            CString NewVal1, NewVal2, NewVal3, NewVal4;
            apply1.GetWindowText(NewVal1);
            apply2.GetWindowText(NewVal2);
            apply3.GetWindowText(NewVal3);
            apply4.GetWindowText(NewVal4);

            // 입력값이 비어 있는 경우 기본값 0으로 설정
            int V1 = NewVal1.IsEmpty() ? 0 : _ttoi(NewVal1);
            int V2 = NewVal2.IsEmpty() ? 0 : _ttoi(NewVal2);
            int V3 = NewVal3.IsEmpty() ? 0 : _ttoi(NewVal3);
            int V4 = NewVal4.IsEmpty() ? 0 : _ttoi(NewVal4);

            // 새로운 값 계산 (B, C, D, E)
            int _NewVal1 = OldVal1 + V1;
            int _NewVal2 = OldVal2 + V2;
            int _NewVal3 = OldVal3 + V3;
            int _NewVal4 = OldVal4 + V4;

            row[1] = CString(std::to_wstring(_NewVal1).c_str()); // B
            row[2] = CString(std::to_wstring(_NewVal2).c_str()); // C
            row[3] = CString(std::to_wstring(_NewVal3).c_str()); // D
            row[4] = CString(std::to_wstring(_NewVal4).c_str()); // E

            // F1 업데이트 (60 ÷ apply5) 및 (100 ÷ apply6) + apply1, apply2 값 반영
            int F1 = OldF1; // 기존 F1 값 유지

            CString App5T, App6T;
            apply5.GetWindowText(App5T);
            apply6.GetWindowText(App6T);

            int Apply5V = _ttoi(App5T);
            int Apply6V = _ttoi(App6T);

            F1 += V1 * 5;
            F1 += V2 * 20;

            if (Apply5V != 0) {
                F1 += 60 / Apply5V;
            }
            if (Apply6V != 0) {
                F1 += 100 / Apply6V;
            }

            // F1 값 누적 업데이트
            row[5] = CString(std::to_wstring(F1).c_str());

            WriteCSVFile(FilePath, data);

            NameFound = true;
            break;
        }
    }

    if (!NameFound) {
        AfxMessageBox(_T("데이터를 찾을 수 없습니다."));
    }
    else {
        AfxMessageBox(_T("적용 완료!"));
    }
}


void CExcelMFCDialog::OnBnClickedButtonSort()
{
    std::vector<std::vector<CString>> data;
    ReadCSVFile(FilePath, data);

    // 데이터 유효성 체크
    if (data.size() <= 1) {
        AfxMessageBox(_T("정렬할 데이터가 부족합니다."));
        return;
    }

    // 헤더 제외한 데이터 정렬
    auto header = data.front();
    std::sort(data.begin() + 1, data.end(), [](const std::vector<CString>& a, const std::vector<CString>& b) {
        if (a.size() < 6 || b.size() < 6) return false;  // 데이터가 부족하면 정렬 안 함

        // F열(합계) 내림차순 정렬
        int sumA = _ttoi(a[5]);
        int sumB = _ttoi(b[5]);
        if (sumA != sumB) {
            return sumA > sumB;  // 합계 기준 정렬
        }

        return a[0].CompareNoCase(b[0]) < 0;  // 이름 가나다순 정렬
        });

    // 정렬된 데이터 다시 파일에 저장
    WriteCSVFile(FilePath, data);

    AfxMessageBox(_T("점수가 높은 순서대로 정렬되었습니다."));
}
