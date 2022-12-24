#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <wincred.h>

#include <taskschd.h>
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")
#pragma warning(disable: 4996)
using namespace std;


int __cdecl wmain(int argc, wchar_t** argv)
{

    wchar_t* str = new wchar_t;
    str[0] = L'rm';
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr))
    {
        printf("\nCoInitializeEx failed: %x", hr);
        return 1;
    }

    hr = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL);

    if (FAILED(hr))
    {
        printf("\nCoInitializeSecurity failed: %x", hr);
        CoUninitialize();
        return 1;
    }

    if (argc == 4) {
        LPCWSTR wszTaskName = argv[1];

        //wstring wstrExecutablePath = _wgetenv(L"WINDIR");
        //wstrExecutablePath += L"\\SYSTEM32\\NOTEPAD.EXE";
        wstring wstrExecutablePath = argv[2];
        ITaskService* pService = NULL;
        hr = CoCreateInstance(CLSID_TaskScheduler,
            NULL,
            CLSCTX_INPROC_SERVER,
            IID_ITaskService,
            (void**)&pService);
        if (FAILED(hr))
        {
            printf("Failed to create an instance of ITaskService: %x", hr);
            CoUninitialize();
            return 1;
        }

        hr = pService->Connect(_variant_t(), _variant_t(),
            _variant_t(), _variant_t());
        if (FAILED(hr))
        {
            printf("ITaskService::Connect failed: %x", hr);
            pService->Release();
            CoUninitialize();
            return 1;
        }

        ITaskFolder* pRootFolder = NULL;
        hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
        if (FAILED(hr))
        {
            printf("Cannot get Root folder pointer: %x", hr);
            pService->Release();
            CoUninitialize();
            return 1;
        }

        pRootFolder->DeleteTask(_bstr_t(wszTaskName), 0);

        ITaskDefinition* pTask = NULL;
        hr = pService->NewTask(0, &pTask);

        pService->Release();
        if (FAILED(hr))
        {
            printf("Failed to CoCreate an instance of the TaskService class: %x", hr);
            pRootFolder->Release();
            CoUninitialize();
            return 1;
        }

        IRegistrationInfo* pRegInfo = NULL;
        hr = pTask->get_RegistrationInfo(&pRegInfo);
        if (FAILED(hr))
        {
            printf("\nCannot get identification pointer: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }
        BSTR Author = SysAllocString(L"Administrator");
        pRegInfo->put_Author(Author);
 /*       pRegInfo->Release();
        if (FAILED(hr))
        {
            printf("\nCannot put identification info: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }*/
        //检查 AppID 证书缓存中是否存在无效或吊销的证书。
        //BSTR Author = SysAllocString(L"检查您"+L" Microsoft 软件保持最新状态。");
        hr = pRegInfo->put_Description(_bstr_t(L"Check the running status of your ")+ _bstr_t(argv[1]) + _bstr_t(L" software."));
        pRegInfo->Release();
        if (FAILED(hr))
        {
            printf("\nCannot put_Description info: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }

        IPrincipal* pPrincipal = NULL;
        hr = pTask->get_Principal(&pPrincipal);
        if (FAILED(hr))
        {
            printf("\nCannot get principal pointer: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }

        hr = pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
        pPrincipal->Release();
        if (FAILED(hr))
        {
            printf("\nCannot put principal info: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }


        ITaskSettings* pSettings = NULL;
        hr = pTask->get_Settings(&pSettings);
        if (FAILED(hr))
        {
            printf("\nCannot get settings pointer: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }

        hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
        pSettings->Release();
        if (FAILED(hr))
        {
            printf("\nCannot put setting information: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }


        IIdleSettings* pIdleSettings = NULL;
        hr = pSettings->get_IdleSettings(&pIdleSettings);
        if (FAILED(hr))
        {
            printf("\nCannot get idle setting information: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }
        BSTR PT5M = SysAllocString(L"PT5M");
        hr = pIdleSettings->put_WaitTimeout(PT5M);
        pIdleSettings->Release();
        if (FAILED(hr))
        {
            printf("\nCannot put idle setting information: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }


        ITriggerCollection* pTriggerCollection = NULL;
        hr = pTask->get_Triggers(&pTriggerCollection);
        if (FAILED(hr))
        {
            printf("\nCannot get trigger collection: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }


        ITrigger* pTrigger = NULL;
        hr = pTriggerCollection->Create(TASK_TRIGGER_TIME, &pTrigger);
        pTriggerCollection->Release();
        if (FAILED(hr))
        {
            printf("\nCannot create trigger: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }

        ITimeTrigger* pTimeTrigger = NULL;
        hr = pTrigger->QueryInterface(
            IID_ITimeTrigger, (void**)&pTimeTrigger);
        pTrigger->Release();
        if (FAILED(hr))
        {
            printf("\nQueryInterface call failed for ITimeTrigger: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }

        hr = pTimeTrigger->put_Id(_bstr_t(L"Trigger1"));
        if (FAILED(hr))
            printf("\nCannot put trigger ID: %x", hr);

        hr = pTimeTrigger->put_EndBoundary(_bstr_t(L"2025-05-02T08:00:00"));
        if (FAILED(hr))
            printf("\nCannot put end boundary on trigger: %x", hr);

        hr = pTimeTrigger->put_StartBoundary(_bstr_t(L"2005-01-01T12:05:00"));
        pTimeTrigger->Release();
        if (FAILED(hr))
        {
            printf("\nCannot add start boundary to trigger: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }

        IRepetitionPattern* IRep = NULL;
        hr = pTimeTrigger->get_Repetition(&IRep);
        IRep->put_Interval(_bstr_t(L"PT") + _bstr_t(argv[3]));
        IRep->Release();
        if (FAILED(hr))
        {
            printf("\nCannot put_Interval: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }



        IActionCollection* pActionCollection = NULL;
        hr = pTask->get_Actions(&pActionCollection);
        if (FAILED(hr))
        {
            printf("\nCannot get Task collection pointer: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }


        IAction* pAction = NULL;
        hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
        pActionCollection->Release();
        if (FAILED(hr))
        {
            printf("\nCannot create the action: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }

        IExecAction* pExecAction = NULL;

        hr = pAction->QueryInterface(
            IID_IExecAction, (void**)&pExecAction);
        pAction->Release();
        if (FAILED(hr))
        {
            printf("\nQueryInterface call failed for IExecAction: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }


        hr = pExecAction->put_Path(_bstr_t(wstrExecutablePath.c_str()));
        pExecAction->Release();
        if (FAILED(hr))
        {
            printf("\nCannot put action path: %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }

        IRegisteredTask* pRegisteredTask = NULL;
        hr = pRootFolder->RegisterTaskDefinition(
            _bstr_t(wszTaskName),
            pTask,
            TASK_CREATE_OR_UPDATE,
            _variant_t(),
            _variant_t(),
            TASK_LOGON_INTERACTIVE_TOKEN,
            _variant_t(L""),
            &pRegisteredTask);
        if (FAILED(hr))
        {
            printf("\nError saving the Task : %x", hr);
            pRootFolder->Release();
            pTask->Release();
            CoUninitialize();
            return 1;
        }

        printf("\n Success! Task successfully registered. ");


        pRootFolder->Release();
        pTask->Release();
        pRegisteredTask->Release();
        CoUninitialize();
        return 0;
        
    }
    else if (_wtoi(argv[1])== 1&& argc == 3)
    {
        ITaskService* pService = NULL;
        hr = CoCreateInstance(CLSID_TaskScheduler,
            NULL,
            CLSCTX_INPROC_SERVER,
            IID_ITaskService,
            (void**)&pService);
        if (FAILED(hr))
        {
            printf("Failed to create an instance of ITaskService: %x", hr);
            CoUninitialize();
            return 1;
        }

        hr = pService->Connect(_variant_t(), _variant_t(),
            _variant_t(), _variant_t());
        if (FAILED(hr))
        {
            printf("ITaskService::Connect failed: %x", hr);
            pService->Release();
            CoUninitialize();
            return 1;
        }

        ITaskFolder* pRootFolder = NULL;
        hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
        if (FAILED(hr))
        {
            printf("Cannot get Root folder pointer: %x", hr);
            pService->Release();
            CoUninitialize();
            return 1;
        }

        pRootFolder->DeleteTask(_bstr_t(argv[2]), 0);
        if (FAILED(hr))
        {
            printf("\nError Delete the Task : %x", hr);
            pRootFolder->Release();
            pRootFolder->Release();
            CoUninitialize();
            return 1;
        }
        else
        {
            printf("Delete Success");
            return 0;
        }
    }
    else
    {
        printf("Create Servers: ITaskServers.exe ServerNmae Hacker.exe 5H\nDelete Servers: ITaskServers.exe 1 ServerNmae");
        return 1;
    }
    
    
}