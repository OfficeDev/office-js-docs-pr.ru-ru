---
title: Очистка кэша Office
description: Узнайте, как очищать кэш Office на компьютере.
ms.date: 01/21/2020
localization_priority: Priority
ms.openlocfilehash: 68e5c022671844ee44bf8ca8ac00bc5af6564bad
ms.sourcegitcommit: 43166612e9b4bf7a73312a572663c8696353dbc6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/29/2020
ms.locfileid: "41580971"
---
# <a name="clear-the-office-cache"></a><span data-ttu-id="efdef-103">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="efdef-103">Clear the Office cache</span></span>

<span data-ttu-id="efdef-104">Можно удалить надстройку, ранее установленную в Windows, на компьютерах Mac или в iOS, очистив кэш Office на компьютере.</span><span class="sxs-lookup"><span data-stu-id="efdef-104">You can remove an add-in that you've previously sideloaded on Windows, Mac, or iOS by clearing the Office cache on your computer.</span></span> 

<span data-ttu-id="efdef-105">Кроме того, если вы изменяете манифест надстройки (например, обновляете имена файлов значков или текст команд надстройки), следует очистить кэш Office, а потом заново установить надстройку с помощью обновленного манифеста.</span><span class="sxs-lookup"><span data-stu-id="efdef-105">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you should clear the Office cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="efdef-106">В этом случае надстройка будет отображаться в Office в соответствии с обновленным манифестом.</span><span class="sxs-lookup"><span data-stu-id="efdef-106">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>

## <a name="clear-the-office-cache-on-windows"></a><span data-ttu-id="efdef-107">Очистка кэша Office в Windows</span><span class="sxs-lookup"><span data-stu-id="efdef-107">Clear the Office cache on Windows</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="efdef-108">Excel, Word и PowerPoint</span><span class="sxs-lookup"><span data-stu-id="efdef-108">Excel, Word, and PowerPoint</span></span> 

<span data-ttu-id="efdef-109">Чтобы очистить кэш Office в Windows для Excel, Word и PowerPoint, удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="efdef-109">To clear the Office cache on Windows for Excel, Word, and PowerPoint, delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

### <a name="outlook-windows-10"></a><span data-ttu-id="efdef-110">Outlook (Windows 10)</span><span class="sxs-lookup"><span data-stu-id="efdef-110">Outlook (Windows 10)</span></span>

<span data-ttu-id="efdef-111">Чтобы очистить кэш Outlook в Windows 10, когда надстройка работает в Microsoft Edge, можно использовать Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="efdef-111">To clear the Outlook cache on Windows 10 when the add-in is running in Microsoft Edge, you can use the Microsoft Edge DevTools.</span></span>

> [!TIP]
> <span data-ttu-id="efdef-112">Если вы хотите только загрузить неопубликованную надстройку, чтобы отразить последние изменения в ее исходных файлах HTML или JavaScript, не нужно использовать описанные ниже действия, чтобы очистить кэш.</span><span class="sxs-lookup"><span data-stu-id="efdef-112">If you're just wanting the sideloaded add-in to reflect recent changes to its HTML or JavaScript source files, you shouldn't need to use the following steps to clear the cache.</span></span> <span data-ttu-id="efdef-113">Вместо этого просто переместите фокус в область задач надстройки (щелкнув в любом месте области задач) и нажмите клавишу **F5**, чтобы перезагрузить надстройку.</span><span class="sxs-lookup"><span data-stu-id="efdef-113">Instead, just put focus in the add-in's task pane (by clicking anywhere within the task pane) and then press **F5** to reload the add-in.</span></span> 

> [!NOTE]
> <span data-ttu-id="efdef-114">Чтобы очистить кэш Outlook с помощью следующих действий, в вашей надстройке должна быть область задач.</span><span class="sxs-lookup"><span data-stu-id="efdef-114">To clear the Outlook cache using the following steps, your add-in must have a task pane.</span></span> <span data-ttu-id="efdef-115">Если в вашей надстройке нет пользовательского интерфейса (например, она использует функцию [проверки при отправке](/outlook/add-ins/outlook-on-send-addins)), потребуется добавить в надстройку область задач, использующую такой же домен для [SourceLocation](../reference/manifest/sourcelocation.md), прежде чем вы сможете использовать указанные ниже действия для очистки кэша.</span><span class="sxs-lookup"><span data-stu-id="efdef-115">If your add-in is a UI-less add-in -- for example, one that uses the [on-send](/outlook/add-ins/outlook-on-send-addins) feature -- you'll need to add a task pane to your add-in that uses the same domain for [SourceLocation](../reference/manifest/sourcelocation.md), before you can use the following steps to clear the cache.</span></span>

1. <span data-ttu-id="efdef-116">Установите [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).</span><span class="sxs-lookup"><span data-stu-id="efdef-116">Install the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj).</span></span>

2. <span data-ttu-id="efdef-117">Откройте свою надстройку в Outlook.</span><span class="sxs-lookup"><span data-stu-id="efdef-117">Open your add-in in Outlook.</span></span>

3. <span data-ttu-id="efdef-118">Запустите Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="efdef-118">Run the Microsoft Edge DevTools.</span></span>

4. <span data-ttu-id="efdef-119">В Microsoft Edge DevTools перейдите на вкладку **Локальные**. Имя вашей надстройки будет указано в списке.</span><span class="sxs-lookup"><span data-stu-id="efdef-119">In the Microsoft Edge DevTools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

5. <span data-ttu-id="efdef-120">Выберите имя надстройки, чтобы присоединить отладчик к надстройке.</span><span class="sxs-lookup"><span data-stu-id="efdef-120">Select the add-in name to attach the debugger to your add-in.</span></span> <span data-ttu-id="efdef-121">Откроется новое окно Microsoft Edge DevTools, когда отладчик присоединяется к надстройке.</span><span class="sxs-lookup"><span data-stu-id="efdef-121">A new Microsoft Edge DevTools window will open when the debugger attaches to your add-in.</span></span>

6. <span data-ttu-id="efdef-122">На вкладке **Сеть** в новом окне нажмите кнопку **Очистить кэш**.</span><span class="sxs-lookup"><span data-stu-id="efdef-122">On the **Network** tab of the new window, select the **Clear cache** button.</span></span>

    ![Снимок экрана Microsoft Edge DevTools с выделенной кнопкой "Очистить кэш"](../images/edge-devtools-clear-cache.png)

7. <span data-ttu-id="efdef-124">Если эти действия не привели к нужному результату, вы также можете нажать кнопку **Всегда обновлять с сервера**.</span><span class="sxs-lookup"><span data-stu-id="efdef-124">If completing these steps doesn't produce the desired result, you can also select the **Always refresh from server** button.</span></span>

    ![Снимок экрана Microsoft Edge DevTools с выделенной кнопкой "Всегда обновлять с сервера"](../images/edge-devtools-refresh-from-server.png)

## <a name="clear-the-office-cache-on-mac"></a><span data-ttu-id="efdef-126">Очистка кэша Office на компьютерах Mac</span><span class="sxs-lookup"><span data-stu-id="efdef-126">Clear the Office cache on Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

##  <a name="clear-the-office-cache-on-ios"></a><span data-ttu-id="efdef-127">Очистка кэша Office в iOS</span><span class="sxs-lookup"><span data-stu-id="efdef-127">Clear the Office cache on iOS</span></span>

<span data-ttu-id="efdef-128">Чтобы очистить кэш Office в iOS, вызовите `window.location.reload(true)` в JavaScript в надстройке, чтобы запустить принудительную перезагрузку.</span><span class="sxs-lookup"><span data-stu-id="efdef-128">To clear the Office cache on iOS, call `window.location.reload(true)` from JavaScript in the add-in to force a reload.</span></span> <span data-ttu-id="efdef-129">Также можно переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="efdef-129">Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="efdef-130">См. также</span><span class="sxs-lookup"><span data-stu-id="efdef-130">See also</span></span>

- [<span data-ttu-id="efdef-131">Отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="efdef-131">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
- [<span data-ttu-id="efdef-132">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="efdef-132">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="efdef-133">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="efdef-133">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="efdef-134">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="efdef-134">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="efdef-135">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="efdef-135">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)

