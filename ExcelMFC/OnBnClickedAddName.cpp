// OnBnClickedAddName.cpp: 구현 파일
//

#include "pch.h"
#include "ExcelMFC.h"
#include "afxdialogex.h"
#include "OnBnClickedAddName.h"
#include <afxrich.h>



// OnBnClickedAddName 대화 상자

IMPLEMENT_DYNAMIC(OnBnClickedAddName, CDialogEx)

OnBnClickedAddName::OnBnClickedAddName(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_EXCELMFC_DIALOG, pParent)
{

}

OnBnClickedAddName::~OnBnClickedAddName()
{
}

void OnBnClickedAddName::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(OnBnClickedAddName, CDialogEx)
	ON_BN_CLICKED(IDC_BUTTON1, &OnBnClickedAddName::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_OnBnClickedBrowse, &OnBnClickedAddName::OnBnClickedOnbnclickedbrowse)
	ON_EN_CHANGE(IDC_RICHEDIT21, &OnBnClickedAddName::OnEnChangeRichedit21)
	ON_EN_CHANGE(IDC_EDIT1, &OnBnClickedAddName::OnEnChangeEdit1)
END_MESSAGE_MAP()


// OnBnClickedAddName 메시지 처리기


void OnBnClickedAddName::OnBnClickedButton1()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
}


void OnBnClickedAddName::OnBnClickedOnbnclickedbrowse()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
}

void OnBnClickedAddName::OnEnChangeRichedit21()
{

}





void OnBnClickedAddName::OnEnChangeEdit1()
{
	// TODO:  RICHEDIT 컨트롤인 경우, 이 컨트롤은
	// CDialogEx::OnInitDialog() 함수를 재지정 
	//하고 마스크에 OR 연산하여 설정된 ENM_CHANGE 플래그를 지정하여 CRichEditCtrl().SetEventMask()를 호출하지 않으면
	// ENM_CHANGE가 있으면 마스크에 ORed를 플래그합니다.

	// TODO:  여기에 컨트롤 알림 처리기 코드를 추가합니다.
}
