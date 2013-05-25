
#pragma once

class CVisioFrameWnd;

class CAddinApp : public CWinApp
{
public:
	CAddinApp();
	~CAddinApp();

	void OnCommand(UINT id);

	Visio::IVApplicationPtr GetVisioApp();
	void SetVisioApp(Visio::IVApplicationPtr app);

	Office::IRibbonUIPtr GetRibbon();
	void SetRibbon(Office::IRibbonUIPtr ribbon);

	virtual BOOL InitInstance();
	virtual int ExitInstance();

	int GetCapturedCount();
	void OnPointCaptured(Visio::IVShapePtr shape, short row);

	UINT GetActiveCommand();

private:
	struct Impl;
	Impl* m_impl;
};

extern CAddinApp theApp;
