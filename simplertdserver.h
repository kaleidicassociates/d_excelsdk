// from http://www.codeproject.com/Articles/245265/Guide-to-Writing-Custom-Functions-in-Excel-Part-II

// SimpleRTDServer.h : Declaration of the CSimpleRTDServer

#pragma once
#include "resource.h"       // main symbols
#include <map>
#include <string>
#include <boost>

#include "RTDExample.h"

using namespace ATL;
using namespace boost;
using namespace std;

// CSimpleRTDServer

class ATL_NO_VTABLE CSimpleRTDServer :
	public CComObjectRootEx<ccomsinglethreadmodel>,
	public CComCoClass<csimplertdserver,>,
	public IDispatchImpl<isimplertdserver, *wmajor="*/" *wminor="*/">,
	public IDispatchImpl<irtdserver, wmajor="*/" wminor="*/">
{
public:
	CSimpleRTDServer()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_SIMPLERTDSERVER)

	DECLARE_NOT_AGGREGATABLE(CSimpleRTDServer)

	BEGIN_COM_MAP(CSimpleRTDServer)
		COM_INTERFACE_ENTRY(ISimpleRTDServer)
		COM_INTERFACE_ENTRY2(IDispatch, IRtdServer)
		COM_INTERFACE_ENTRY(IRtdServer)
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:


	// IRtdServer Methods
public:
	STDMETHODIMP ServerStart(IRTDUpdateEvent * CallbackObject, long * pfRes);
	STDMETHODIMP ConnectData(long TopicID, SAFEARRAY * * Strings, VARIANT_BOOL * GetNewValues, VARIANT * pvarOut);
	STDMETHODIMP RefreshData(long * TopicCount, SAFEARRAY * * parrayOut);
	STDMETHODIMP DisconnectData(long TopicID);
	STDMETHODIMP Heartbeat(long * pfRes);
	STDMETHODIMP ServerTerminate();

private:
	std::list<std::wstring> StringsAsList(SAFEARRAY * * Strings);
	IRTDUpdateEvent * m_callBackObj;
	std::map<long,>> m_results;
	std::list<long> m_new_results;
	boost::thread m_backgroundThread;
};
class CWorkerTask
{
private:
	long m_topicID;
	IRTDUpdateEvent * m_callBackObj;
	std::list<long> * m_pNewResults;
	std::list<std::wstring> m_args;
public:
	CWorkerTask(long topicID, IRTDUpdateEvent * callBackObj, std::list<long> * newResults, std::list<std::wstring> args);
	double operator()();
};
OBJECT_ENTRY_AUTO(__uuidof(SimpleRTDServer), CSimpleRTDServer)
