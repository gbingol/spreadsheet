#pragma once

#include <string>
#include <tuple>
#include <optional>

#include <wx/wx.h>
#include <wx/grid.h>
#include <wx/xml/xml.h>

#include "dllimpexp.h"

namespace grid
{
	class CWorksheetBase;
	class Cell;

	DLLGRID std::string ColNumtoLetters(size_t num);

	DLLGRID int LetterstoColumnNumber(const std::string& letters);

	DLLGRID wxFont StringtoFont(const wxString& str);
	DLLGRID std::wstring FonttoString(const wxFont& font);


	//tabs and newlines
	DLLGRID wxString GenerateTabString(
		const CWorksheetBase* ws,
		const wxGridCellCoords& TopLeft,
		const wxGridCellCoords& BottomRight);

	DLLGRID std::pair<wxGridCellCoords, wxGridCellCoords>
		AddSelToClipbrd(const CWorksheetBase* ws);

	

	/****************************************************** */


	DLLGRID std::optional<wxXmlDocument> CreateXMLDoc(const wxString& XMLString);

	DLLGRID wxString GetXMLData();

	DLLGRID bool SupportsXML();

	DLLGRID std::vector<Cell> XMLDocToCells(
		const grid::CWorksheetBase* ws,
		const wxXmlDocument& xmlDoc);


	DLLGRID std::tuple<long, long, bool>
		XMLNodeToRowsCols(const wxXmlNode* xmlNode);

	DLLGRID bool ParseXMLDoc(
		grid::CWorksheetBase* ws,
		const wxXmlDocument& xmlDoc);

	//UTF8 string
	DLLGRID std::string GenerateXMLString(
		const grid::CWorksheetBase* ws,
		const wxGridCellCoords& TopLeft,
		const wxGridCellCoords& BottomRight);

	//UTF8 string
	DLLGRID std::string GenerateXMLString(grid::CWorksheetBase* ws);


	//used by worksheet and others (probably they should not directly use it!)
	struct XMLDataFormat : public wxDataFormat
	{
		XMLDataFormat() : wxDataFormat("XMLDataFormat") {}
	};


	class XMLDataObject : public wxTextDataObject
	{
	public:
		XMLDataObject(const std::string& xmlUTF8 = "") :
			wxTextDataObject(xmlUTF8)
		{
			SetFormat(XMLDataFormat());
		}
	};
}