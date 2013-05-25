// AddIn.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"
#include "Addin.h"

#include "AddIn_i.h"
#include "AddIn_i.c"

#include "lib/Visio.h"

CComModule _Module;

// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
	return _Module.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _Module.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
	// registers object, typelib and all interfaces in typelib
	HRESULT hr = _Module.DllRegisterServer();
	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _Module.DllUnregisterServer();
	return hr;
}

STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
	HRESULT hr = E_FAIL;
	// MSVC will call "regsvr32 /i:user" if "per-user registration" is set as a
	// linker option - so handle that here (its also handle for anyone else to
	// be able to manually install just for themselves.)
	static const wchar_t szUserSwitch[] = L"user";
	if (pszCmdLine != NULL)
	{
		if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
		{
			AtlSetPerUserRegistration(true);
			// But ATL still barfs if you try and register a COM category, so
			// just arrange to not do that.
			_AtlComModule.m_ppAutoObjMapFirst = _AtlComModule.m_ppAutoObjMapLast;
		}
	}
	if (bInstall)
	{
		hr = DllRegisterServer();
		if (FAILED(hr))
		{
			DllUnregisterServer();
		}
	}
	else
	{
		hr = DllUnregisterServer();
	}
	return hr;
}

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

BOOL CAddinApp::InitInstance()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// Initialize COM stuff
	if (FAILED(_Module.Init(ObjectMap, AfxGetInstanceHandle(), &LIBID_AddinLib)))
		return FALSE;

	return TRUE;
}

int CAddinApp::ExitInstance() 
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	_Module.Term();

	return CWinApp::ExitInstance();
}

struct CAddinApp::Impl
{
	Impl()
	{
		captured = 0;
		command = 0;
	}

	Visio::IVApplicationPtr app;
	Office::IRibbonUIPtr	ribbon;

	struct Point
	{
		Visio::IVShapePtr shape;
		short row;

		double x;
		double y;

		bstr_t formula_x;
		bstr_t formula_y;
		bstr_t formula_dir_x;
		bstr_t formula_dir_y;
	};

	Point points[2];

	UINT command;
	int captured;

	void OnClick(UINT cmd_id)
	{
		while (captured > 0)
			RemovePoint();

		if (cmd_id == command)
		{
			app->DoCmd(Visio::visCmdDRPointerTool);
			command = 0;
		}
		else
		{
			if (command == 0)
				app->DoCmd(Visio::visCmdDRConnectionTool);

			command = cmd_id;
		}

		ribbon->Invalidate();
	}

	void RemovePoint()
	{
		if (captured > 0)
		{
			Point& point = points[--captured];

			point.shape->DeleteRow(Visio::visSectionConnectionPts, point.row);
			point.shape = NULL;
		}
	}

	void AddPoint(Point point)
	{
		short rows = point.shape->GetRowCount(Visio::visSectionConnectionPts);
		point.row = point.shape->AddRow(Visio::visSectionConnectionPts, rows, 0);

		Visio::IVCellPtr cell_x = point.shape->CellsSRC[Visio::visSectionConnectionPts][point.row][0];
		cell_x->FormulaForceU = point.formula_x;

		Visio::IVCellPtr cell_y = point.shape->CellsSRC[Visio::visSectionConnectionPts][point.row][1];
		cell_y->FormulaForceU = point.formula_y;

		Visio::IVCellPtr cell_dir_x = point.shape->CellsSRC[Visio::visSectionConnectionPts][point.row][2];
		cell_dir_x->FormulaForceU = point.formula_dir_x;

		Visio::IVCellPtr cell_dir_y = point.shape->CellsSRC[Visio::visSectionConnectionPts][point.row][3];
		cell_dir_y->FormulaForceU = point.formula_dir_y;

		captured = 1;
		points[0] = point;
	}

	void Execute()
	{
		double dx = points[1].x - points[0].x;
		double dy = points[1].y - points[0].y;

		Visio::IVWindowPtr window;
		if (FAILED(app->get_ActiveWindow(&window)) || window == NULL)
			return;

		Visio::IVSelectionPtr selection;
		if (FAILED(window->get_Selection(&selection)) || selection == NULL)
			return;

		switch (command)
		{

		case ID_TwoPointsMove:
			{
				RemovePoint();
				RemovePoint();

				selection->Move(dx, dy);
				break;
			}

		case ID_TwoPointsCopy:
			{
				RemovePoint();

				Point saved = points[0];
				RemovePoint();

				double x1, y1, x2, y2;
				selection->BoundingBox(Visio::visBBoxUprightWH, &x1, &y1, &x2, &y2);

				selection->Duplicate();

				Visio::IVSelectionPtr new_selection;
				if (FAILED(window->get_Selection(&new_selection)) || new_selection == NULL)
					return;

				double dup_x1, dup_y1, dup_x2, dup_y2;
				new_selection->BoundingBox(Visio::visBBoxUprightWH, &dup_x1, &dup_y1, &dup_x2, &dup_y2);

				dx -= (dup_x1 - x1);
				dy -= (dup_y1 - y1);

				new_selection->Move(dx, dy);

				AddPoint(saved);

				window->Selection = selection;
				break;
			}
		}
	}
};

void CAddinApp::OnCommand(UINT id)
{
	switch (id)
	{
	case ID_About:
		{
			CDialog dlg(IDD_DIALOG1);
			dlg.DoModal();
			break;
		}

	default:
		{
			m_impl->OnClick(id);
			break;
		}
	}
}

Visio::IVApplicationPtr CAddinApp::GetVisioApp()
{
	return m_impl->app;
}

void CAddinApp::SetVisioApp( Visio::IVApplicationPtr app )
{
	m_impl->app = app;
}

Office::IRibbonUIPtr CAddinApp::GetRibbon()
{
	return m_impl->ribbon;
}

void CAddinApp::SetRibbon(Office::IRibbonUIPtr ribbon)
{
	m_impl->ribbon = ribbon;
}

CAddinApp::CAddinApp()
{
	m_impl = new Impl();
}

CAddinApp::~CAddinApp()
{
	delete m_impl;
}

UINT CAddinApp::GetActiveCommand()
{
	return m_impl->command;
}

int CAddinApp::GetCapturedCount()
{
	return m_impl->captured;
}

void CAddinApp::OnPointCaptured(Visio::IVShapePtr shape, short row )
{
	Impl::Point& point = m_impl->points[m_impl->captured];

	point.shape = shape;
	point.row = row;

	Visio::IVCellPtr cell_x = shape->CellsSRC[Visio::visSectionConnectionPts][row][0];
	double x = cell_x->ResultIU;
	point.formula_x = cell_x->FormulaU;

	Visio::IVCellPtr cell_y = shape->CellsSRC[Visio::visSectionConnectionPts][row][1];
	double y = cell_y->ResultIU;
	point.formula_y = cell_y->FormulaU;

	Visio::IVCellPtr cell_dir_x = shape->CellsSRC[Visio::visSectionConnectionPts][row][2];
	point.formula_dir_x = cell_dir_x->FormulaU;

	Visio::IVCellPtr cell_dir_y = shape->CellsSRC[Visio::visSectionConnectionPts][row][3];
	point.formula_dir_y = cell_dir_y->FormulaU;

	shape->XYToPage(x, y, &point.x, &point.y);

	++m_impl->captured;

	if (m_impl->captured == 2)
	{
		m_impl->Execute();
	}

	m_impl->ribbon->Invalidate();
}

CAddinApp theApp;
