#pragma once
#include "afxdialogex.h"


// OnBnClickedAddName 대화 상자

class OnBnClickedAddName : public CDialogEx
{
	DECLARE_DYNAMIC(OnBnClickedAddName)

public:
	OnBnClickedAddName(CWnd* pParent = nullptr);   // 표준 생성자입니다.
	virtual ~OnBnClickedAddName();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXCELMFC_DIALOG };
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 지원입니다.

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedOnbnclickedbrowse();
	afx_msg void OnEnChangeRichedit21();
	afx_msg void OnEnChangeEdit1();
	CComboBox m_combobox1;
	CEdit m_fix_test;
};
