---
title: Загрузка неопубликованных надстроек Office в Office в Интернете для тестирования
description: Протестируйте надстройку Office в Office в Интернете для загрузки неопубликованных приложений.
ms.date: 09/21/2020
localization_priority: Normal
ms.openlocfilehash: 709461d19fbf4602db3ba5bd9c40f495d0dbbd52
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175537"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a><span data-ttu-id="5fac7-103">Загрузка неопубликованных надстроек Office в Office в Интернете для тестирования</span><span class="sxs-lookup"><span data-stu-id="5fac7-103">Sideload Office Add-ins in Office on the web for testing</span></span>

<span data-ttu-id="5fac7-104">Загрузка неопубликованной надстройки Office позволит быстро установить ее для тестирования, не размещая в каталоге надстроек.</span><span class="sxs-lookup"><span data-stu-id="5fac7-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="5fac7-105">Для загрузки неопубликованных приложений можно использовать как Microsoft 365, так и Office в Интернете.</span><span class="sxs-lookup"><span data-stu-id="5fac7-105">Sideloading can be done in either Microsoft 365 or Office on the web.</span></span> <span data-ttu-id="5fac7-106">Эта процедура слегка различается для каждой из двух платформ.</span><span class="sxs-lookup"><span data-stu-id="5fac7-106">The procedure is slightly different for the two platforms.</span></span>

<span data-ttu-id="5fac7-107">При загрузке неопубликованной надстройки ее манифест хранится в локальном хранилище браузера. Поэтому если очистить кэш браузера или поменять браузер, процедуру придется повторить.</span><span class="sxs-lookup"><span data-stu-id="5fac7-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>

> [!NOTE]
> <span data-ttu-id="5fac7-p102">Загрузка неопубликованных надстроек, описанная в этой статье, поддерживается в Word, Excel и PowerPoint. Соответствующие действия касательно надстройки Outlook приведены в статье [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="5fac7-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

<span data-ttu-id="5fac7-110">В видео ниже показано, как загрузить неопубликованную надстройку в Office в Интернете или классическое приложение.</span><span class="sxs-lookup"><span data-stu-id="5fac7-110">The following video walks you through the process of sideloading your add-in in Office on the web or desktop.</span></span>

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a><span data-ttu-id="5fac7-111">Загрузка неопубликованной надстройки Office в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="5fac7-111">Sideload an Office Add-in in Office on the web</span></span>

1. <span data-ttu-id="5fac7-112">Откройте [Office в Интернете](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="5fac7-112">Open [Office on the web](https://office.live.com/).</span></span>

2. <span data-ttu-id="5fac7-113">В разделе Начало **работы с веб-приложениями**выберите **Excel**, **Word**или **PowerPoint**; а затем откройте новый документ.</span><span class="sxs-lookup"><span data-stu-id="5fac7-113">In **Get started with the online apps now**, choose **Excel**, **Word**, or **PowerPoint**; and then open a new document.</span></span>

3. <span data-ttu-id="5fac7-114">Откройте вкладку **Вставка** на ленте и в разделе **надстройки** выберите надстройки **Office**.</span><span class="sxs-lookup"><span data-stu-id="5fac7-114">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>

4. <span data-ttu-id="5fac7-115">В диалоговом окне **надстройки Office** откройте вкладку **Мои** надстройки, выберите **Управление моими**надстройками, а затем **отправьте надстройку**.</span><span class="sxs-lookup"><span data-stu-id="5fac7-115">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>

    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="5fac7-117">**Найдите** файл манифеста надстройки и выберите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="5fac7-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>

    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)

6. <span data-ttu-id="5fac7-p103">Убедитесь, что надстройка установлена. Например, если надстройка вызывается командой, эта команда должна появиться на ленте или в контекстном меню. Если же у вас надстройка области задач, должна появиться область.</span><span class="sxs-lookup"><span data-stu-id="5fac7-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
> <span data-ttu-id="5fac7-122">Чтобы протестировать надстройку Office с помощью Microsoft EDGE, необходимо выполнить дополнительные действия по настройке.</span><span class="sxs-lookup"><span data-stu-id="5fac7-122">To test your Office Add-in with Microsoft Edge, an additional configuration step is required.</span></span> <span data-ttu-id="5fac7-123">В командной строке Windows выполните следующую строку: `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`</span><span class="sxs-lookup"><span data-stu-id="5fac7-123">In a Windows Command Prompt, run the following line: `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`</span></span>

## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="5fac7-124">Загрузка неопубликованной надстройки Office в Office 365</span><span class="sxs-lookup"><span data-stu-id="5fac7-124">Sideload an Office Add-in in Office 365</span></span>

1. <span data-ttu-id="5fac7-125">Войдите в свою учетную запись Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="5fac7-125">Sign in to your Microsoft 365 account.</span></span>

2. <span data-ttu-id="5fac7-126">Откройте средство запуска приложений в левой части панели инструментов и выберите **Excel**, **Word**или **PowerPoint**, а затем создайте новый документ.</span><span class="sxs-lookup"><span data-stu-id="5fac7-126">Open the App Launcher on the left end of the toolbar and select **Excel**, **Word**, or **PowerPoint**, and then create a new document.</span></span>

3. <span data-ttu-id="5fac7-127">Действия 3–6 совпадают с действиями в предыдущем разделе **Загрузка неопубликованной надстройки Office в Office в Интернете**.</span><span class="sxs-lookup"><span data-stu-id="5fac7-127">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office on the web**.</span></span>

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="5fac7-128">Загрузка неопубликованной настройки при использовании Visual Studio</span><span class="sxs-lookup"><span data-stu-id="5fac7-128">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="5fac7-129">Если вы разрабатываете надстройки с помощью Visual Studio, процесс загрузки неопубликованной надстройки будет аналогичным.</span><span class="sxs-lookup"><span data-stu-id="5fac7-129">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="5fac7-130">Единственное различие состоит в том, что необходимо обновить значение элемента **SourceURL** в манифесте, чтобы включить в него полный URL-адрес расположения, в котором развернута надстройка.</span><span class="sxs-lookup"><span data-stu-id="5fac7-130">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="5fac7-131">Хотя неопубликованные надстройки можно загружать из Visual Studio в Office в Интернете, их невозможно отлаживать из Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="5fac7-131">Although you can sideload add-ins from Visual Studio to Office on the web, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="5fac7-132">Для этого вам потребуются средства отладки браузера.</span><span class="sxs-lookup"><span data-stu-id="5fac7-132">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="5fac7-133">Дополнительные сведения см. в статье [Отладка надстроек в Office в Интернете](debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="5fac7-133">For more information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="5fac7-134">В Visual Studio откройте окно **Свойства**, выбрав **Вид** -> **Окно свойств**.</span><span class="sxs-lookup"><span data-stu-id="5fac7-134">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="5fac7-135">В **обозревателе решений** выберите веб-проект.</span><span class="sxs-lookup"><span data-stu-id="5fac7-135">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="5fac7-136">В окне **Свойства** отобразятся свойства проекта.</span><span class="sxs-lookup"><span data-stu-id="5fac7-136">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="5fac7-137">В окне "Свойства" скопируйте **URL-адрес SSL**.</span><span class="sxs-lookup"><span data-stu-id="5fac7-137">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="5fac7-138">В проекте надстройки откройте XML-файл манифеста.</span><span class="sxs-lookup"><span data-stu-id="5fac7-138">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="5fac7-139">Убедитесь, что вы изменяете исходный XML-файл.</span><span class="sxs-lookup"><span data-stu-id="5fac7-139">Be sure you are editing the source XML.</span></span> <span data-ttu-id="5fac7-140">Для проектов некоторых типов в Visual Studio откроется визуальное представление XML-файла, которое не будет работать на следующем шаге.</span><span class="sxs-lookup"><span data-stu-id="5fac7-140">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="5fac7-141">Найдите и замените все экземпляры **~remoteAppUrl/** только что скопированным URL-адресом SSL.</span><span class="sxs-lookup"><span data-stu-id="5fac7-141">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="5fac7-142">В зависимости от типа проекта отобразится несколько вариантов замены, и появятся новые URL-адреса, похожие на `https://localhost:44300/Home.html`.</span><span class="sxs-lookup"><span data-stu-id="5fac7-142">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="5fac7-143">Сохраните XML-файл.</span><span class="sxs-lookup"><span data-stu-id="5fac7-143">Save the XML file.</span></span>
7. <span data-ttu-id="5fac7-144">Щелкните веб-проект правой кнопкой мыши и выберите **Отладка** -> **Запустить новый экземпляр**.</span><span class="sxs-lookup"><span data-stu-id="5fac7-144">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="5fac7-145">Веб-проект будет выполнен без запуска Office.</span><span class="sxs-lookup"><span data-stu-id="5fac7-145">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="5fac7-146">В Office в Интернете загрузите неопубликованную надстройку согласно инструкциям, приведенным выше в разделе [Загрузка неопубликованной надстройки Office в Office в Интернете](#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="5fac7-146">From Office on the web, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office on the web](#sideload-an-office-add-in-in-office-on-the-web).</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="5fac7-147">Удаление надстройки неопубликованные</span><span class="sxs-lookup"><span data-stu-id="5fac7-147">Remove a sideloaded add-in</span></span>

<span data-ttu-id="5fac7-148">Вы можете удалить ранее созданную надстройку неопубликованные, очистив кэш браузера.</span><span class="sxs-lookup"><span data-stu-id="5fac7-148">You can remove a previously sideloaded add-in by clearing your browser's cache.</span></span> <span data-ttu-id="5fac7-149">Кроме того, при внесении изменений в манифест надстройки (например, для обновления имен файлов значков или текста команд надстройки) может потребоваться очистить кэш, а затем повторно Загрузка неопубликованных надстройку с помощью обновленного манифеста.</span><span class="sxs-lookup"><span data-stu-id="5fac7-149">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you may need to clear the cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="5fac7-150">В этом случае надстройка будет отображаться в Office в соответствии с обновленным манифестом.</span><span class="sxs-lookup"><span data-stu-id="5fac7-150">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>
