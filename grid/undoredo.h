#pragma once

#include <string>
#include <vector>

#include <wx/wx.h>
#include <wx/grid.h>

#include "ws_cell.h"

#include "dllimpexp.h"

namespace grid
{
	class CWorksheetBase;

	class DLLGRID WorksheetUndoRedoEvent
	{
	public:
		WorksheetUndoRedoEvent(CWorksheetBase* worksheet, bool CanRedo);
		virtual ~WorksheetUndoRedoEvent();

		void ShowWorksheet(); //Show the worksheet where events are happening

		auto GetEventSource() const {
			return m_WSBase;
		}

		bool CanRedo() const {
			return m_CanRedo;
		}

		virtual std::wstring GetToolTip(bool IsUndo) = 0;
		virtual void Undo() = 0;
		virtual void Redo() = 0;

	protected:
		/*
		Initially say a cell has a value of 0 and then we changed it to 10
		Undo should restore the cell to the initial value of 0
		Redo should restore the cell to the last value of 10
		*/
		CWorksheetBase* m_WSBase;

		bool m_CanRedo;
	};




	//triggered when data is changed by user by typing
	class DLLGRID CellDataChanged : public WorksheetUndoRedoEvent
	{
	public:
		CellDataChanged(CWorksheetBase* worksheet, int row, int col);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		std::wstring m_InitVal; //Before the change
		std::wstring m_LastVal; //After the change

	private:
		int m_row;
		int m_col;
	};



	//triggered with SetCellValue
	class DLLGRID CellValueChangedEvent : public WorksheetUndoRedoEvent
	{
	public:
		CellValueChangedEvent(CWorksheetBase* worksheet, int row, int col);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		std::wstring m_InitVal; //Before the change
		std::wstring m_LastVal; //After the change

	private:
		int m_row;
		int m_col;
	};



	class DLLGRID CellContentDeleted : public WorksheetUndoRedoEvent
	{
	public:
		CellContentDeleted(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		void SetInitialCells(const std::vector<Cell>& Cells);

		void SetCoords(const wxGridCellCoords& TL, const wxGridCellCoords& BR) {
			m_TL = TL;
			m_BR = BR;
		}

	private:
		std::vector<Cell> m_InitVal; //Before the change

		wxGridCellCoords m_TL;
		wxGridCellCoords m_BR;
	};



	class DLLGRID CellBGColorChangedEvent : public WorksheetUndoRedoEvent
	{
	public:
		CellBGColorChangedEvent(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		std::vector<Cell> m_InitVal; //Before the change
		std::vector<Cell> m_LastVal; //After the change

		wxGridCellCoords m_TL;
		wxGridCellCoords m_BR;
	};



	class DLLGRID TextColorChangedEvent : public WorksheetUndoRedoEvent
	{
	public:
		TextColorChangedEvent(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		std::vector<Cell> m_InitVal; //Before the change
		std::vector<Cell> m_LastVal; //After the change

		wxGridCellCoords m_TL;
		wxGridCellCoords m_BR;
	};



	class DLLGRID FontChangedEvent : public WorksheetUndoRedoEvent
	{
	public:
		FontChangedEvent(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		std::vector<Cell> m_InitVal;
		std::vector<Cell> m_LastVal;

		std::string m_FontProperty; //Info on changed property, such as "Size", "Face", "Bold" ...

		wxGridCellCoords m_TL;
		wxGridCellCoords m_BR;
	};



	class DLLGRID CellAlignmentChangedEvent : public WorksheetUndoRedoEvent
	{
	public:
		CellAlignmentChangedEvent(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		std::vector<Cell> m_InitVal;
		std::vector<Cell> m_LastVal;

		wxGridCellCoords m_TL;
		wxGridCellCoords m_BR;
	};



	class DLLGRID RowsDeleted : public WorksheetUndoRedoEvent
	{
	public:
		RowsDeleted(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		void SetInfo(int Start, int Length) 
		{
			m_StartPos = Start;
			m_Length = Length;
		}

		void SetInitialCells(const std::vector<Cell>& InitCells);

	private:
		std::vector<Cell> m_InitVal; //Before the change

		int m_StartPos;
		int m_Length;
	};



	class DLLGRID RowsInserted : public WorksheetUndoRedoEvent
	{
	public:
		RowsInserted(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		void SetInfo(int Start, int Length)
		{
			m_StartPos = Start;
			m_Length = Length;
		}

	private:
		int m_StartPos;
		int m_Length;
	};





	class DLLGRID ColumnsDeleted : public WorksheetUndoRedoEvent
	{
	public:
		ColumnsDeleted(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		void SetInfo(int Start, int Length)
		{
			m_StartPos = Start;
			m_Length = Length;
		}

		void SetInitialCells(const std::vector<Cell>& InitCells);

	private:

		std::vector<Cell> m_InitVal; //Before the change

		int m_StartPos;
		int m_Length;
	};





	class DLLGRID ColumnsInserted : public WorksheetUndoRedoEvent
	{
	public:
		ColumnsInserted(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		void SetInfo(int Start, int Length) 
		{
			m_StartPos = Start;
			m_Length = Length;
		}

	private:
		int m_StartPos;
		int m_Length;
	};



	class DLLGRID RowColSizeChanged : public WorksheetUndoRedoEvent
	{

	public:
		enum class ENTITY { ROW = 0, COL };

		RowColSizeChanged(CWorksheetBase* worksheet, ENTITY entity, int Number);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		int m_PrevSize;
		int m_FinalSize;

	private:
		int m_ColRowNumber;
		ENTITY m_Entity;
	};


	class DLLGRID DataPasted : public WorksheetUndoRedoEvent
	{
	public:
		DataPasted(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		void SetCoords(const wxGridCellCoords& TL, const wxGridCellCoords& BR) 
		{
			m_TL = TL;
			m_BR = BR;
		}

		//first:TL, second:BR
		void SetCoords(const std::pair< wxGridCellCoords, wxGridCellCoords>& Coords) 
		{
			m_TL = Coords.first;
			m_BR = Coords.second;
		}

		void SetPaste(int pastewhat) 
		{
			m_PasteWhat = pastewhat;
		}

	private:
		int m_PasteWhat;

		wxGridCellCoords m_TL;
		wxGridCellCoords m_BR;

	};



	class DLLGRID DataCut : public WorksheetUndoRedoEvent
	{
	public:
		DataCut(CWorksheetBase* worksheet);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		void SetCells(const std::vector<Cell>& Cells);

		void SetCoords(const wxGridCellCoords& TL, const wxGridCellCoords& BR) 
		{
			m_TL = TL;
			m_BR = BR;
		}

	private:

		std::vector<Cell>	m_Value;

		wxGridCellCoords m_TL;
		wxGridCellCoords m_BR;
	};



	class DLLGRID DataMovedEvent : public WorksheetUndoRedoEvent
	{
	public:
		//can also be copied during moving (Ctrl button)
		DataMovedEvent(CWorksheetBase* worksheet, bool Moved = true);

		void Undo();
		void Redo();

		std::wstring GetToolTip(bool IsUndo) override;

		bool operator==(const WorksheetUndoRedoEvent& other) const;

		void SetInitCoords(const wxGridCellCoords& TL, const wxGridCellCoords& BR) 
		{
			m_Init_TL = TL;
			m_Init_BR = BR;
		}

		void SetFinalCoords(const wxGridCellCoords& TL, const wxGridCellCoords& BR) 
		{
			m_Final_TL = TL;
			m_Final_BR = BR;
		}

		void SetCells(const std::vector<Cell>& Cells);

	private:

		std::vector<Cell>	m_CellValues; //Initial location and values

		wxGridCellCoords m_Init_TL, m_Init_BR;
		wxGridCellCoords m_Final_TL, m_Final_BR;


		//if user presses Ctrl button then it is copied
		bool m_Moved = true;
	};


}
