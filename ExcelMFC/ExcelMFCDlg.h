#pragma once
#include <vector>
#include <afx.h>
#include <afxrich.h>
class CExcelMFCDialog : public CDialogEx
{
public:
    CExcelMFCDialog(CWnd* pParent = nullptr);  // 생성자
    virtual ~CExcelMFCDialog();               // 소멸자 추가

#ifdef AFX_DESIGN_TIME
    enum { IDD = IDD_EXCELMFC_DIALOG };
#endif

protected:
    virtual void DoDataExchange(CDataExchange* pDX);

protected:
    DECLARE_MESSAGE_MAP()

public:
    CString SName;       // 이름 입력 변수
    CString FilePath;   // 선택한 파일 경로
    void ReadCSVFile(const CString& filePath, std::vector<std::vector<CString>>& data);  // CSV 읽기
    void WriteCSVFile(const CString& filePath, const std::vector<std::vector<CString>>& data);  // CSV 쓰기
    void BackupCSVFile(const CString& filePath);

    afx_msg void OnBnClickedAddName();  // 이름 추가 버튼 핸들러
    afx_msg void OnBnClickedBrowse();  // 파일 열기 버튼 핸들러
    afx_msg void OnEnChangeEdit1();
    afx_msg void OnBnClickButton3();
    afx_msg void OnBnClickedBackup();
    std::vector<std::vector<CString>> data;
    CComboBox m_combobox1;
    CEdit m_editvalue4;
    CEdit m_editvalue5;
    CEdit m_editvalue6;
    CEdit m_editvalue8;
    CEdit m_editvalue7;
    afx_msg void OnBnClickedApply(); // Apply 버튼 클릭 시
    CEdit apply1, apply2, apply3, apply4, apply5, apply6;
    afx_msg void OnBnClickedButtonSort();
};