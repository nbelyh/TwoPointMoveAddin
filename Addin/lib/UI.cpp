
#include "stdafx.h"

#include "Language.h"
#include "Addin.h"
#include "Utils.h"
#include "TextFile.h"
#include "UI.h"
#include "PictureConvert.h"

#define ADDON_NAME	L"TwoPoints"

_ATL_FUNC_INFO ClickEventInfo = { CC_STDCALL, VT_EMPTY, 2, { VT_DISPATCH, VT_BOOL|VT_BYREF } };


ClickEventRedirector::ClickEventRedirector(IUnknownPtr punk) 
	: m_punk(punk)
{
	DispEventAdvise(punk);
}

ClickEventRedirector::~ClickEventRedirector()
{
	DispEventUnadvise(m_punk);
}

void __stdcall ClickEventRedirector::OnClick(IDispatch* pButton, VARIANT_BOOL* pCancel)
{
	try
	{
		Office::_CommandBarButtonPtr button;
		pButton->QueryInterface(__uuidof(Office::_CommandBarButton), (void**)&button);

		CComBSTR parameter;
		button->get_Parameter(&parameter);

		UINT cmd_id = StrToInt(parameter);

		if (cmd_id > 0)
			theApp.OnCommand(cmd_id);

	}
	catch (_com_error)
	{
		// 
	}
}

void AddinUi::FillMenuItems(Office::CommandBarControlsPtr menu_items, CMenu* popup_menu )
{
	// For each items in the menu,
	bool begin_group = false;
	for (UINT i = 0; i < popup_menu->GetMenuItemCount(); ++i)
	{
		CMenu* sub_menu = popup_menu->GetSubMenu(i);

		// set item caption
		CString item_caption;
		popup_menu->GetMenuString(i, item_caption, MF_BYPOSITION);

		// create new menu item.
		Office::CommandBarControlPtr item;
		menu_items->Add(
			variant_t(long(Office::msoControlButton)), 
			vtMissing, 
			vtMissing, 
			vtMissing, 
			variant_t(true),
			&item);

		// obtain command id from menu
		UINT command_id = popup_menu->GetMenuItemID(i);

		CString caption;
		caption.LoadString(command_id);
		item->put_Caption(bstr_t(caption));

		CString parameter;
		parameter.Format(L"%d", command_id);
		item->put_Parameter(bstr_t(parameter));

		// Set unique tag, so that the command is not lost
		CString tag;
		tag.Format(L"%s_%d", ADDON_NAME, command_id);
		item->put_Tag(bstr_t(tag));

		CBitmap bm_picture;
		// if we are command button
		Office::MsoControlType item_type;
		if (SUCCEEDED(item->get_Type(&item_type)) && item_type == Office::msoControlButton && command_id > 0)
		{
			IPictureDispPtr picture;
			IPictureDispPtr mask;
			if (SUCCEEDED(CustomUiGetPng(MAKEINTRESOURCE(command_id), &picture, &mask)))
			{
				try
				{
					// cast to button. hopefully we will always succeed here
					Office::_CommandBarButtonPtr button = item;

					button->put_Picture(picture);
					button->put_Mask(mask);
				}
				catch (_com_error)
				{
					// There Some problems; hope to never get here.
				}
			}
		}

		m_buttons.Add(new ClickEventRedirector(item));

	}
}

void AddinUi::CreateCommandBarsMenu(Visio::IVApplicationPtr app)
{
	LanguageLock lock(GetAppLanguage(app));

	Office::_CommandBarsPtr cbs = app->CommandBars;

	CMenu menu;
	menu.LoadMenu(IDR_MENU);

	Office::CommandBarPtr cb;
	if (SUCCEEDED(cbs->get_Item(variant_t(ADDON_NAME), &cb)))
	{
		Office::CommandBarControlsPtr controls;
		cb->get_Controls(&controls);

		FillMenuItems(controls, menu.GetSubMenu(0));
	}
}

void AddinUi::DestroyCommandBarsMenu()
{
	for (size_t i = 0; i < m_buttons.GetCount(); ++i)
		delete m_buttons[i];

	m_buttons.RemoveAll();
}

HRESULT GetRibbonText(BSTR * RibbonXml)
{
	LPBYTE content = NULL;
	DWORD content_length = 0;

	ASSERT_RETURN_VALUE(LoadResourceFromModule(
		AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_RESOURCE_H), L"TEXT", &content, &content_length), S_FALSE);

	CMemFile mf(content, content_length);
	CTextFileRead rdr(&mf);

	CMapStringToString replacements;

	CString line;
	while (rdr.ReadLine(line))
	{
		CSimpleArray<CString> tokens;

		CString token;
		token.Preallocate(30);

		for (LPCWSTR pos = line; *pos; ++pos)
		{
			if (*pos == ' ' || *pos == '\t')
			{
				if (!token.IsEmpty())
				{
					tokens.Add(token);
					token = CString();
				}
			}
			else
			{
				token += *pos;
			}
		}

		if (!token.IsEmpty())
			tokens.Add(token);

		if (tokens.GetSize() != 3)
			continue;

		if (tokens[0] != "#define")
			continue;

		replacements[tokens[1]] = tokens[2];
	}

	CString ribbon = LoadTextFromModule(AfxGetInstanceHandle(), IDR_RIBBON);

	for (int pos = 0; pos < ribbon.GetLength(); ++pos)
	{
		if (ribbon[pos] != '{')
			continue;

		int endpos = ribbon.Find('}', pos);
		ASSERT_CONTINUE(endpos != -1);

		CString token = ribbon.Mid(pos+1, endpos-pos-1);
		CString token_found;

		ASSERT_CONTINUE(replacements.Lookup(token, token_found));	

		ribbon.Delete(pos, endpos-pos+1);
		ribbon.Insert(pos, token_found);

		pos += (token_found.GetLength()-1);
	}

	*RibbonXml = ribbon.AllocSysString();
	return S_OK;
}
