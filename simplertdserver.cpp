// from http://www.codeproject.com/Articles/245265/Guide-to-Writing-Custom-Functions-in-Excel-Part-II

// need to call regsvr32 RTDExample.dll to register

// SimpleRTDServer.cpp : Implementation of CSimpleRTDServer

#include "stdafx.h"
#include "SimpleRTDServer.h"
#include <boost>
#include <boost>
#include <string>
#include <list>

using namespace std;
using namespace boost;

// CSimpleRTDServer

// Called during loading of the dll. CallbackObject is what we use to notify Excel that
// we're done calculating some values and we're ready to refresh the data in the spreadsheet.
HRESULT CSimpleRTDServer::ServerStart(IRTDUpdateEvent * CallbackObject, long * pfRes /* <= 0 means failure */)
{
	if(CallbackObject == NULL || pfRes == NULL)
	{
		return E_POINTER;
	}
	m_callBackObj = CallbackObject;
	*pfRes = 1;
	return S_OK;
}
// Whenever a new topic is needed Excel will call this. The GetNewValues parameter tells Excel to use
// the previous value until the call to RefreshData or display the default while waiting.
HRESULT CSimpleRTDServer::ConnectData(long TopicID, SAFEARRAY * * Strings, VARIANT_BOOL * GetNewValues, VARIANT * pvarOut)
{
	CWorkerTask worker(TopicID, m_callBackObj, &m_new_results, StringsAsList(Strings));
	boost::packaged_task<double> pt(worker);
	
	
	m_results[TopicID] = boost::move(pt.get_future());
	boost::thread task(boost::move(pt)); // start thread.
	task.detach();
	return S_OK;
}

// After we call UpdateNotify, Excel calls this function to request the values
// of the topics that have been updated.
HRESULT CSimpleRTDServer::RefreshData(long * TopicCount, SAFEARRAY * * parrayOut)
{
	HRESULT hr = S_OK;
	if(TopicCount == NULL || parrayOut == NULL || (*parrayOut != NULL))
	{
		hr = E_POINTER;
	}
	else
	{
		*TopicCount = m_new_results.size();
		SAFEARRAYBOUND bounds[2];
		VARIANT value;
		long index[2];

		// Create a safe array		
		bounds[0].cElements = 2;
		bounds[0].lLbound = 0;
		bounds[1].cElements = *TopicCount;
		bounds[1].lLbound = 0;
		*parrayOut = SafeArrayCreate(VT_VARIANT, 2, bounds);
		int i = 0;
		for(list<long>::const_iterator itor = m_new_results.begin(); itor != m_new_results.end(); ++itor)
		{
			index[0] = 0;
			index[1] = i;
			
			VariantInit(&value);
			value.vt = VT_I4;
			value.lVal = *itor; // Topic ID
			SafeArrayPutElement(*parrayOut, index, &value);

			index[0] = 1;
			VariantInit(&value);
			value.vt = VT_R8;
			if(!m_results[*itor].is_ready())
				m_results[*itor].wait();
			value.dblVal = m_results[*itor].get(); // Result
			SafeArrayPutElement(*parrayOut, index, &value);			
		}
		m_new_results.clear();
		hr = S_OK;
	}
	return hr;
}
// Excel tells us that it doesn't need a topic by calling this.
// We remove the TopicID from the results. In a real add-in you'd
// probably want to check if the thread is still running and stop it.
HRESULT CSimpleRTDServer::DisconnectData(long TopicID)
{
	m_results.erase(TopicID);
	return S_OK;
}
// Excel calls this to determine if we're still alive.
// If pfRes is non-negative then we're still good.
HRESULT CSimpleRTDServer::Heartbeat(long * pfRes)
{
	HRESULT hr = S_OK;
	if(pfRes == NULL)
		hr = E_POINTER;
	else
		*pfRes = 1;
	return hr;
}
// Before Excel unloads the dll it calls this.
HRESULT CSimpleRTDServer::ServerTerminate()
{
	return S_OK;
}
// Converts the parameters supplied to ConnectData "Strings" into a list of wstrings 
// to make it easier to work with.
std::list<std::wstring> CSimpleRTDServer::StringsAsList(SAFEARRAY * * Strings)
{
	std::list<std::wstring> result;
	LONG lbound, ubound;
	SafeArrayGetLBound(*Strings,1,&lbound);
	SafeArrayGetUBound(*Strings,1,&ubound);
	
	VARIANT* pvar;
	SafeArrayAccessData(*Strings, (void HUGEP**) &pvar);
	for(long i = lbound; i <= ubound; i++)
	{
		BSTR bs = pvar[i].bstrVal;
		
		result.push_back(std::wstring(bs,SysStringLen(bs)));
	}
	SafeArrayUnaccessData(*Strings);

	return result;
}

// CWorkerTask

// This is what we'll be doing in the thread. Typically you might want to query a service of somekind here.
double CWorkerTask::operator()()
{
	this_thread::sleep(boost::posix_time::seconds(5));
	double dPerimeter = 0.0;
	for(std::list<std::wstring>::const_iterator itor = m_args.begin(); itor != m_args.end(); ++itor)
	{
		dPerimeter += boost::lexical_cast<double>(*itor);
	}
	m_pNewResults->push_front(m_topicID);
	m_callBackObj->UpdateNotify();
	
	return dPerimeter;
}
CWorkerTask::CWorkerTask(long topicID, IRTDUpdateEvent *callBackObj,std::list<long> * newResults, std::list<std::wstring> args)
{
	m_callBackObj = callBackObj;
	m_pNewResults = newResults;
	m_topicID = topicID;
	m_args = args;
}
