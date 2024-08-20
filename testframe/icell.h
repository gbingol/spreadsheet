#pragma once

#include <vector>



#include <wx/wx.h>
#include <wx/aui/auibook.h>
#include <wx/aui/aui.h>
#include <wx/popupwin.h>
#include <wx/tglbtn.h>
#include <wx/colordlg.h>

#include <grid/worksheetbase.h>
#include <grid/ntbkbase.h>
#include <grid/workbookbase.h>




namespace ICELL
{
	class CWorksheetNtbk;
	class CWorkbook;


	class CWorksheet :public grid::CWorksheetBase
	{
	public:
		CWorksheet(wxWindow* panel,
			CWorkbook* workbook,
			wxWindowID id,
			const wxPoint& pos,
			const wxSize& size,
			long style,
			wxString WindowName,
			int nrows = 1000,
			int ncols = 100);

		virtual ~CWorksheet();

		CWorkbook* GetWorkbook() const { 
			return m_Workbook;
		}

		wxMenu* GetContextMenu() const
		{
			return m_ContextMenu;
		}

	protected:
		void OnRightClick(wxGridEvent& event);


	private:
		CWorkbook* m_Workbook = nullptr;
		wxMenu* m_ContextMenu{nullptr};
		wxWindow* m_ParentWnd;
	};


	/********************************************************* */

	class CWorksheetNtbk : public grid::CWorksheetNtbkBase
	{

	public:
		CWorksheetNtbk(CWorkbook* parent);
		virtual ~CWorksheetNtbk();

		grid::CWorksheetBase* CreateWorksheet(wxWindow* wnd, const std::wstring& Label, int nrows, int ncols) const;

		auto GetContextMenu() const {
			return m_ContextMenu;
		}

	private:
		CWorkbook* m_Workbook{nullptr};
	};



	/************************************************************** */


	class CWorkbook : public grid::CWorkbookBase
	{
	public:
		CWorkbook(wxWindow* parent);
		virtual ~CWorkbook();

		CWorksheetNtbk* GetWorksheetNotebook() const override
		{
			return (CWorksheetNtbk*)m_WSNtbk;
		}

		bool AddNewWorksheet(const std::wstring& tblname = L"", int nrows = 1000, int ncols = 50) override
		{
			return GetWorksheetNotebook()->AddNewWorksheet(tblname, nrows, ncols);
		}

		
	};
}