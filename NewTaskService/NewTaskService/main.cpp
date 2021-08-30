#include <Windows.h>
#include <taskschd.h>

#include <stdio.h>

#pragma comment(lib, "taskschd.lib")

int delete_task(ITaskFolder *pfolder, const WCHAR *taskname)
{
    VARIANT vtProp;
    HRESULT hr;
    int result = 0;

    //VariantInit(&vtProp);
    //vtProp.vt = VT_BSTR;
    //vtProp.bstrVal = taskname;
    hr = pfolder->DeleteTask((BSTR)taskname, 0);
    if (SUCCEEDED(hr))
    {
        result = 1;
    }
    return(result);
}

int task_schedule_initialize(ITaskService **pservice, ITaskFolder **pfolder)
{
    VARIANT servername, user, domain, password;
    HRESULT hr;
    int result = 0;

    *pservice = NULL;
    *pfolder = NULL;

    // 创建一个任务服务（Task Service）实例
    hr = CoCreateInstance(CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (LPVOID*)pservice);
    if (SUCCEEDED(hr))
    {
        VariantInit(&servername);
        VariantInit(&user);
        VariantInit(&domain);
        VariantInit(&password);

        hr = (*pservice)->Connect(servername, user, domain, password);
        if (SUCCEEDED(hr))
        {
            // 获取 Root Task Folder 的指针，这个指针指向的是新注册的任务
            hr = (*pservice)->GetFolder((BSTR)L"\\", pfolder);
            if (SUCCEEDED(hr))
            {
                result = 1;
            }
        }
    }
    return(result);
}
int task_schedule_new(ITaskService *pservice, ITaskFolder *pfolder, const WCHAR *taskname, const WCHAR *path, const WCHAR *parameters, const WCHAR *author)
{
    VARIANT userid, password, sddl;
    ITaskDefinition *ptaskdefinition = NULL;
    IRegistrationInfo *preginfo = NULL;
    IPrincipal *pprincipal = NULL;
    ITaskSettings *psetting = NULL;
    IActionCollection *pactioncollect = NULL;
    IAction *paction = NULL;
    IExecAction *pexecaction = NULL;
    ITriggerCollection *ptriggers = NULL;
    ITrigger *ptrigger = NULL;
    IRegisteredTask *pregisteredtask = NULL;
    HRESULT hr;
    int result = 0;

    // 如果存在相同的计划任务，则删除
    delete_task(pfolder, taskname);

    // 创建任务定义对象来创建任务
    hr = pservice->NewTask(0, &ptaskdefinition);
    if (SUCCEEDED(hr))
    {
        hr = ptaskdefinition->get_RegistrationInfo(&preginfo);
        if (SUCCEEDED(hr))
        {
            // 设置作者信息
            hr = preginfo->put_Author((BSTR)author);
            preginfo->Release();
            hr = ptaskdefinition->get_Principal(&pprincipal);
            if (SUCCEEDED(hr))
            {
                // 设置登录类型
                hr = pprincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
                // 设置运行权限
                // 最高权限
                hr = pprincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);
                pprincipal->Release();
                hr = ptaskdefinition->get_Settings(&psetting);
                if (SUCCEEDED(hr))
                {
                    // 设置其他信息
                    hr = psetting->put_StopIfGoingOnBatteries(VARIANT_FALSE);
                    hr = psetting->put_DisallowStartIfOnBatteries(VARIANT_FALSE);
                    hr = psetting->put_AllowDemandStart(VARIANT_TRUE);
                    hr = psetting->put_StartWhenAvailable(VARIANT_FALSE);
                    hr = psetting->put_MultipleInstances(TASK_INSTANCES_PARALLEL);
                    psetting->Release();

                    // 创建执行动作
                    hr = ptaskdefinition->get_Actions(&pactioncollect);
                    if (SUCCEEDED(hr))
                    {
                        // 创建执行操作
                        hr = pactioncollect->Create(TASK_ACTION_EXEC, &paction);
                        pactioncollect->Release();
                        if (SUCCEEDED(hr))
                        {
                            // 设置执行程序路径和参数
                            hr = paction->QueryInterface(IID_IExecAction, (PVOID *)&pexecaction);
                            paction->Release();
                            if (SUCCEEDED(hr))
                            {
                                // 设置程序路径和参数
                                hr = pexecaction->put_Path((BSTR)path);
                                if (SUCCEEDED(hr))
                                {
                                    wprintf(L"%s\r\n", path);
                                }
                                pexecaction->put_Arguments((BSTR)parameters);
                                pexecaction->Release();

                                // 创建触发器，实现用户登陆自启动
                                hr = ptaskdefinition->get_Triggers(&ptriggers);
                                if (SUCCEEDED(hr))
                                {
                                    hr = ptriggers->Create(TASK_TRIGGER_LOGON, &ptrigger);
                                    if (SUCCEEDED(hr))
                                    {
                                        VariantInit(&userid);
                                        VariantInit(&password);
                                        VariantInit(&sddl);

                                        // 注册任务计划
                                        hr = pfolder->RegisterTaskDefinition((BSTR)taskname,
                                            ptaskdefinition,
                                            TASK_CREATE_OR_UPDATE,
                                            userid,
                                            password,
                                            TASK_LOGON_INTERACTIVE_TOKEN,
                                            sddl,
                                            &pregisteredtask);
                                        if (SUCCEEDED(hr))
                                        {
                                            pregisteredtask->Release();

                                            result = 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        ptaskdefinition->Release();
    }

    return(result);
}

int wmain(int argc, WCHAR *argv[])
{
    ITaskService *pservice;
    ITaskFolder *pfolder;
    HRESULT hr;
    int flag;

    if (argc > 1)
    {
        // 初始化COM
        hr = CoInitialize(NULL);
        if (SUCCEEDED(hr))
        {
            if (task_schedule_initialize(&pservice, &pfolder))
            {
                flag = task_schedule_new(pservice, pfolder, L"520", argv[1], L"", L"");

                getchar();

                flag = delete_task(pfolder, L"520");

                getchar();
            }

            CoUninitialize();
        }
    }

	return(0);
}