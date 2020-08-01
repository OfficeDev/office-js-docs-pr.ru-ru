---
title: Очистка кэша Office
description: Узнайте, как очищать кэш Office на компьютере.
ms.date: 05/22/2020
localization_priority: Normal
ms.openlocfilehash: fe8184a5bd084450c3be6d3d13707bf92f5c508f
ms.sourcegitcommit: 7d5407d3900d2ad1feae79a4bc038afe50568be0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2020
ms.locfileid: "46530494"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="2da97-103">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="2da97-103">Clear the Office cache</span></span>

<span data-ttu-id="2da97-104">Можно удалить надстройку, ранее установленную в Windows, на компьютерах Mac или в iOS, очистив кэш Office на компьютере.</span><span class="sxs-lookup"><span data-stu-id="2da97-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span>

<span data-ttu-id="2da97-105">Кроме того, если вы изменяете манифест надстройки (например, обновляете имена файлов значков или текст команд надстройки), следует очистить кэш Office, а потом заново установить надстройку с помощью обновленного манифеста.</span><span class="sxs-lookup"><span data-stu-id="2da97-105">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="2da97-106">В этом случае надстройка будет отображаться в Office в соответствии с обновленным манифестом.</span><span class="sxs-lookup"><span data-stu-id="2da97-106">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="2da97-107">Очистка кэша Office в Windows</span><span class="sxs-lookup"><span data-stu-id="2da97-107">Clear the Office cache on Windows</span></span>

<span data-ttu-id="2da97-108">Чтобы удалить все надстройки неопубликованные из Excel, Word и PowerPoint, удалите содержимое папки:</span><span class="sxs-lookup"><span data-stu-id="2da97-108">To remove all sideloaded add-ins from Excel, Word, and PowerPoint, delete the contents of the folder:</span></span>

```
%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
```

<span data-ttu-id="2da97-109">Если следующая папка существует, удалите ее содержимое:</span><span class="sxs-lookup"><span data-stu-id="2da97-109">If the following folder exists, delete its contents too:</span></span>

```
%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

<span data-ttu-id="2da97-110">Чтобы удалить надстройку неопубликованные из Outlook, выполните действия, описанные в статье [Загрузка неопубликованных Outlook Add-ins for test](../outlook/sideload-outlook-add-ins-for-testing.md) , чтобы найти надстройку в разделе " **Настраиваемые** надстройки" диалогового окна со списком установленных надстроек. Нажмите кнопку с многоточием ( `...` ) для надстройки, а затем выберите **Удалить** , чтобы удалить эту надстройку.</span><span class="sxs-lookup"><span data-stu-id="2da97-110">To remove a sideloaded add-in from Outlook, use the steps outlined in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in and then choose **Remove** to remove that specific add-in.</span></span> <span data-ttu-id="2da97-111">Если эта надстройка не работает, удалите содержимое `Wef` папки, как указано выше для Excel, Word и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="2da97-111">If this add-in removal doesn't work, then delete the contents of the `Wef` folder as noted previously for Excel, Word, and PowerPoint.</span></span>

<span data-ttu-id="2da97-112">Чтобы очистить кэш в Office на Windows 10, когда надстройка работает в Microsoft Edge, вы можете использовать Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="2da97-112">Additionally, to clear the Office cache on Windows 10 when the add-in is running in Microsoft Edge, you can use the Microsoft Edge DevTools.</span></span>

> [!TIP]
> <span data-ttu-id="2da97-113">Если вы хотите, чтобы надстройка неопубликованные отражала последние изменения в исходных файлах HTML или JavaScript, очистка кэша не требуется.</span><span class="sxs-lookup"><span data-stu-id="2da97-113">If you only want the sideloaded add-in to reflect recent changes to its HTML or JavaScript source files, you shouldn't need to clear the cache.</span></span> <span data-ttu-id="2da97-114">Вместо этого просто переместите фокус в область задач надстройки (щелкнув в любом месте области задач) и нажмите клавишу **F5**, чтобы перезагрузить надстройку.</span><span class="sxs-lookup"><span data-stu-id="2da97-114">Instead, just put focus in the add-in's task pane (by clicking anywhere within the task pane) and then press **F5** to reload the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="2da97-115">Чтобы очистить кэш Outlook с помощью следующих действий, в вашей надстройке должна быть панель задач.</span><span class="sxs-lookup"><span data-stu-id="2da97-115">To clear the Office cache using the following steps, your add-in must have a task pane.</span></span> <span data-ttu-id="2da97-116">Если в вашей надстройке нет пользовательского интерфейса (например, она использует функцию [проверки при отправке](../outlook/outlook-on-send-addins.md)), потребуется добавить в надстройку область задач, использующую такой же домен для [SourceLocation](../reference/manifest/sourcelocation.md), прежде чем вы сможете использовать указанные ниже действия для очистки кэша.</span><span class="sxs-lookup"><span data-stu-id="2da97-116">If your add-in is a UI-less add-in -- for example, one that uses the [on-send](../outlook/outlook-on-send-addins.md) feature -- you'll need to add a task pane to your add-in that uses the same domain for [SourceLocation](../reference/manifest/sourcelocation.md), before you can use the following steps to clear the cache.</span></span>

1. <span data-ttu-id="2da97-117">Установите [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).</span><span class="sxs-lookup"><span data-stu-id="2da97-117">Install the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).</span></span>

2. <span data-ttu-id="2da97-118">Откройте надстройку в клиенте Office.</span><span class="sxs-lookup"><span data-stu-id="2da97-118">Open your add-in in the Office client.</span></span>

3. <span data-ttu-id="2da97-119">Запустите Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="2da97-119">Run the Microsoft Edge DevTools.</span></span>

4. <span data-ttu-id="2da97-120">В Microsoft Edge DevTools перейдите на вкладку **Локальные**. Имя вашей надстройки будет указано в списке.</span><span class="sxs-lookup"><span data-stu-id="2da97-120">In the Microsoft Edge DevTools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

5. <span data-ttu-id="2da97-121">Выберите имя надстройки, чтобы присоединить отладчик к надстройке.</span><span class="sxs-lookup"><span data-stu-id="2da97-121">Select the add-in name to attach the debugger to your add-in.</span></span> <span data-ttu-id="2da97-122">Откроется новое окно Microsoft Edge DevTools, когда отладчик присоединяется к надстройке.</span><span class="sxs-lookup"><span data-stu-id="2da97-122">A new Microsoft Edge DevTools window will open when the debugger attaches to your add-in.</span></span>

6. <span data-ttu-id="2da97-123">На вкладке **Сеть** в новом окне нажмите кнопку **Очистить кэш**.</span><span class="sxs-lookup"><span data-stu-id="2da97-123">On the **Network** tab of the new window, select the **Clear cache** button.</span></span>

    ![Снимок экрана Microsoft Edge DevTools с выделенной кнопкой "Очистить кэш"](../images/edge-devtools-clear-cache.png)

7. <span data-ttu-id="2da97-125">Если эти действия не привели к нужному результату, вы также можете нажать кнопку **Всегда обновлять с сервера**.</span><span class="sxs-lookup"><span data-stu-id="2da97-125">If completing these steps doesn't produce the desired result, you can also select the **Always refresh from server** button.</span></span>

    ![Снимок экрана Microsoft Edge DevTools с выделенной кнопкой "Всегда обновлять с сервера"](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="2da97-127">Очистка кэша Office на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="2da97-127">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="2da97-128">Очистка кэша Office в iOS</span><span class="sxs-lookup"><span data-stu-id="2da97-128">Clear the Office cache on iOS</span></span>

<span data-ttu-id="2da97-129">Чтобы очистить кэш Office в iOS, вызовите `window.location.reload(true)` в JavaScript в надстройке, чтобы запустить принудительную перезагрузку.</span><span class="sxs-lookup"><span data-stu-id="2da97-129">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="2da97-130">Также можно переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="2da97-130">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="2da97-131">См. также</span><span class="sxs-lookup"><span data-stu-id="2da97-131">See also</span></span>

- [<span data-ttu-id="2da97-132">Отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="2da97-132">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [<span data-ttu-id="2da97-133">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="2da97-133">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="2da97-134">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="2da97-134">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="2da97-135">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="2da97-135">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="2da97-136">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="2da97-136">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
