---
title: Загрузка неопубликованных надстроек Office в Office в Интернете для тестирования
description: Тестирование неопубликованной надстройки Office в Office в Интернете путем ее загрузки
ms.date: 02/18/2020
localization_priority: Normal
ms.openlocfilehash: 869cabec737c39d7dded04fe7c52011347e0f314
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163586"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a><span data-ttu-id="91a76-103">Загрузка неопубликованных надстроек Office в Office в Интернете для тестирования</span><span class="sxs-lookup"><span data-stu-id="91a76-103">Sideload Office Add-ins in Office on the web for testing</span></span>

<span data-ttu-id="91a76-104">Загрузка неопубликованной надстройки Office позволит быстро установить ее для тестирования, не размещая в каталоге надстроек.</span><span class="sxs-lookup"><span data-stu-id="91a76-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="91a76-105">Загрузить неопубликованную надстройку можно в Office 365 или Office в Интернете.</span><span class="sxs-lookup"><span data-stu-id="91a76-105">Sideloading can be done in either Office 365 or Office on the web.</span></span> <span data-ttu-id="91a76-106">Эта процедура слегка различается для каждой из двух платформ.</span><span class="sxs-lookup"><span data-stu-id="91a76-106">The procedure is slightly different for the two platforms.</span></span>

<span data-ttu-id="91a76-107">При загрузке неопубликованной надстройки ее манифест хранится в локальном хранилище браузера. Поэтому если очистить кэш браузера или поменять браузер, процедуру придется повторить.</span><span class="sxs-lookup"><span data-stu-id="91a76-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>

> [!NOTE]
> <span data-ttu-id="91a76-p102">Загрузка неопубликованных надстроек, описанная в этой статье, поддерживается в Word, Excel и PowerPoint. Соответствующие действия касательно надстройки Outlook приведены в статье [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="91a76-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

<span data-ttu-id="91a76-110">В видео ниже показано, как загрузить неопубликованную надстройку в Office в Интернете или классическое приложение.</span><span class="sxs-lookup"><span data-stu-id="91a76-110">The following video walks you through the process of sideloading your add-in in Office on the web or desktop.</span></span>

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a><span data-ttu-id="91a76-111">Загрузка неопубликованной надстройки Office в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="91a76-111">Sideload an Office Add-in in Office on the web</span></span>

1. <span data-ttu-id="91a76-112">Откройте [Microsoft Office в Интернете](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="91a76-112">Open [Microsoft Office on the web](https://office.live.com/).</span></span>

2. <span data-ttu-id="91a76-113">В разделе  **Начало работы с веб-приложениями** выберите **Excel**,  **Word** или **PowerPoint** и откройте новый документ.</span><span class="sxs-lookup"><span data-stu-id="91a76-113">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>

3. <span data-ttu-id="91a76-114">Откройте вкладку  **Вставка** на ленте и в разделе **Надстройки** выберите **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="91a76-114">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>

4. <span data-ttu-id="91a76-115">В диалоговом окне **Надстройки Office** откройте вкладку **МОИ НАДСТРОЙКИ** и выберите **Управление моими надстройками** > **Отправить надстройку**.</span><span class="sxs-lookup"><span data-stu-id="91a76-115">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>

    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="91a76-117">**Найдите** файл манифеста надстройки и выберите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="91a76-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>

    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)

6. <span data-ttu-id="91a76-p103">Убедитесь, что надстройка установлена. Например, если надстройка вызывается командой, эта команда должна появиться на ленте или в контекстном меню. Если же у вас надстройка области задач, должна появиться область.</span><span class="sxs-lookup"><span data-stu-id="91a76-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="91a76-122">Чтобы протестировать надстройку Office с помощью Microsoft Edge, требуется выполнить два действия по настройке:</span><span class="sxs-lookup"><span data-stu-id="91a76-122">To test your Office Add-in with Microsoft Edge, two configuration steps are required:</span></span> 
>
> - <span data-ttu-id="91a76-123">В командной строке Windows выполните следующую строку: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span><span class="sxs-lookup"><span data-stu-id="91a76-123">In a Windows Command Prompt, run the following line: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span></span>
>
> - <span data-ttu-id="91a76-124">Введите **about:flags** в панели поиска Microsoft Edge, чтобы открыть раздел "Параметры разработчика".</span><span class="sxs-lookup"><span data-stu-id="91a76-124">Enter “**about:flags**” in the Microsoft Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="91a76-125">Установите флажок **Разрешить замыкание на себя для localhost** и перезапустите Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="91a76-125">Check the “**Allow localhost loopback**” option and restart Microsoft Edge.</span></span>

>    ![Параметр "Разрешить замыкание на себя для localhost" в Microsoft Edge с установленным флажком.](../images/allow-localhost-loopback.png)

## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="91a76-127">Загрузка неопубликованной надстройки Office в Office 365</span><span class="sxs-lookup"><span data-stu-id="91a76-127">Sideload an Office Add-in in Office 365</span></span>

1. <span data-ttu-id="91a76-128">Войдите в свою учетную запись Office 365.</span><span class="sxs-lookup"><span data-stu-id="91a76-128">Sign in to your Office 365 account.</span></span>

2. <span data-ttu-id="91a76-129">Откройте средство запуска приложений в левой части панели инструментов, выберите **Excel**, **Word** или **PowerPoint** и создайте документ.</span><span class="sxs-lookup"><span data-stu-id="91a76-129">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>

3. <span data-ttu-id="91a76-130">Действия 3–6 совпадают с действиями в предыдущем разделе **Загрузка неопубликованной надстройки Office в Office в Интернете**.</span><span class="sxs-lookup"><span data-stu-id="91a76-130">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office on the web**.</span></span>

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="91a76-131">Загрузка неопубликованной настройки при использовании Visual Studio</span><span class="sxs-lookup"><span data-stu-id="91a76-131">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="91a76-132">Если вы разрабатываете надстройки с помощью Visual Studio, процесс загрузки неопубликованной надстройки будет аналогичным.</span><span class="sxs-lookup"><span data-stu-id="91a76-132">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="91a76-133">Единственное различие состоит в том, что необходимо обновить значение элемента **SourceURL** в манифесте, чтобы включить в него полный URL-адрес расположения, в котором развернута надстройка.</span><span class="sxs-lookup"><span data-stu-id="91a76-133">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="91a76-134">Хотя неопубликованные надстройки можно загружать из Visual Studio в Office в Интернете, их невозможно отлаживать из Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="91a76-134">Although you can sideload add-ins from Visual Studio to Office on the web, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="91a76-135">Для этого вам потребуются средства отладки браузера.</span><span class="sxs-lookup"><span data-stu-id="91a76-135">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="91a76-136">Дополнительные сведения см. в статье [Отладка надстроек в Office в Интернете](debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="91a76-136">For more information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="91a76-137">В Visual Studio откройте окно **Свойства**, выбрав **Вид** -> **Окно свойств**.</span><span class="sxs-lookup"><span data-stu-id="91a76-137">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="91a76-138">В **обозревателе решений** выберите веб-проект.</span><span class="sxs-lookup"><span data-stu-id="91a76-138">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="91a76-139">В окне **Свойства** отобразятся свойства проекта.</span><span class="sxs-lookup"><span data-stu-id="91a76-139">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="91a76-140">В окне "Свойства" скопируйте **URL-адрес SSL**.</span><span class="sxs-lookup"><span data-stu-id="91a76-140">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="91a76-141">В проекте надстройки откройте XML-файл манифеста.</span><span class="sxs-lookup"><span data-stu-id="91a76-141">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="91a76-142">Убедитесь, что вы изменяете исходный XML-файл.</span><span class="sxs-lookup"><span data-stu-id="91a76-142">Be sure you are editing the source XML.</span></span> <span data-ttu-id="91a76-143">Для проектов некоторых типов в Visual Studio откроется визуальное представление XML-файла, которое не будет работать на следующем шаге.</span><span class="sxs-lookup"><span data-stu-id="91a76-143">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="91a76-144">Найдите и замените все экземпляры **~remoteAppUrl/** только что скопированным URL-адресом SSL.</span><span class="sxs-lookup"><span data-stu-id="91a76-144">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="91a76-145">В зависимости от типа проекта отобразится несколько вариантов замены, и появятся новые URL-адреса, похожие на `https://localhost:44300/Home.html`.</span><span class="sxs-lookup"><span data-stu-id="91a76-145">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="91a76-146">Сохраните XML-файл.</span><span class="sxs-lookup"><span data-stu-id="91a76-146">Save the XML file.</span></span>
7. <span data-ttu-id="91a76-147">Щелкните веб-проект правой кнопкой мыши и выберите **Отладка** -> **Запустить новый экземпляр**.</span><span class="sxs-lookup"><span data-stu-id="91a76-147">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="91a76-148">Веб-проект будет выполнен без запуска Office.</span><span class="sxs-lookup"><span data-stu-id="91a76-148">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="91a76-149">В Office в Интернете загрузите неопубликованную надстройку согласно инструкциям, приведенным выше в разделе [Загрузка неопубликованной надстройки Office в Office в Интернете](#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="91a76-149">From Office on the web, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office on the web](#sideload-an-office-add-in-in-office-on-the-web).</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="91a76-150">Удаление надстройки неопубликованные</span><span class="sxs-lookup"><span data-stu-id="91a76-150">Remove a sideloaded add-in</span></span>

<span data-ttu-id="91a76-151">Вы можете удалить ранее созданную надстройку неопубликованные, очистив кэш браузера.</span><span class="sxs-lookup"><span data-stu-id="91a76-151">You can remove a previously sideloaded add-in by clearing your browser's cache.</span></span> <span data-ttu-id="91a76-152">Кроме того, при внесении изменений в манифест надстройки (например, для обновления имен файлов значков или текста команд надстройки) может потребоваться очистить кэш, а затем повторно Загрузка неопубликованных надстройку с помощью обновленного манифеста.</span><span class="sxs-lookup"><span data-stu-id="91a76-152">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you may need to clear the cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="91a76-153">В этом случае надстройка будет отображаться в Office в соответствии с обновленным манифестом.</span><span class="sxs-lookup"><span data-stu-id="91a76-153">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>
