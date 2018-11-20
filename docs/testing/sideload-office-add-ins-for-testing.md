---
title: Загрузка неопубликованных надстроек Office в Office Online для тестирования
description: Тестирование неопубликованной надстройки Office в Office Online путем ее загрузки
ms.date: 10/19/2018
ms.openlocfilehash: 94138cd0a22f053a9471bf905b8d0838dead15cf
ms.sourcegitcommit: 3a808cf39cbc77056968d53a5957462371ad83a1
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2018
ms.locfileid: "25911230"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a><span data-ttu-id="8b210-103">Загрузка неопубликованных надстроек Office в Office Online для тестирования</span><span class="sxs-lookup"><span data-stu-id="8b210-103">Sideload Office Add-ins in Office Online for testing</span></span>

<span data-ttu-id="8b210-104">Загрузка неопубликованной надстройки Office позволит быстро установить ее для тестирования, не размещая в каталоге надстроек.</span><span class="sxs-lookup"><span data-stu-id="8b210-104">You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog.</span></span> <span data-ttu-id="8b210-105">Загрузить неопубликованную надстройку можно в Office 365 или Office Online.</span><span class="sxs-lookup"><span data-stu-id="8b210-105">Sideloading can be done in either Office 365 or Office Online.</span></span> <span data-ttu-id="8b210-106">Эта процедура слегка различается для каждой из двух платформ.</span><span class="sxs-lookup"><span data-stu-id="8b210-106">The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="8b210-107">При загрузке неопубликованной надстройки ее манифест хранится в локальном хранилище браузера. Поэтому если очистить кэш браузера или поменять браузер, процедуру придется повторить.</span><span class="sxs-lookup"><span data-stu-id="8b210-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="8b210-p102">Загрузка неопубликованных надстроек, описанная в этой статье, поддерживается в Word, Excel и PowerPoint. Соответствующие действия касательно надстройки Outlook приведены в статье [Загрузка неопубликованных надстроек Outlook для тестирования](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="8b210-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="8b210-110">В видео ниже показано, как загрузить неопубликованную надстройку в классическое приложение Office или Office Online.</span><span class="sxs-lookup"><span data-stu-id="8b210-110">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="8b210-111">Загрузка неопубликованной надстройки Office в Office 365</span><span class="sxs-lookup"><span data-stu-id="8b210-111">Sideload an Office Add-in on Office 365</span></span>


1. <span data-ttu-id="8b210-112">Войдите в свою учетную запись Office 365.</span><span class="sxs-lookup"><span data-stu-id="8b210-112">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="8b210-113">Откройте средство запуска приложений в левой части панели инструментов, выберите **Excel**, **Word** или **PowerPoint** и создайте документ.</span><span class="sxs-lookup"><span data-stu-id="8b210-113">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="8b210-114">Откройте вкладку  **Вставка** на ленте и в разделе **Надстройки** выберите **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="8b210-114">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="8b210-115">В диалоговом окне **Надстройки Office** откройте вкладку **МОЯ ОРГАНИЗАЦИЯ** и выберите **Отправить надстройку**.</span><span class="sxs-lookup"><span data-stu-id="8b210-115">On the  **Office Add-ins** dialog, select the **MY ORGANIZATION** tab, and then **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройка Office" со ссылкой "Отправить надстройку" в верхнем левом углу](../images/office-add-ins.png)

5.  <span data-ttu-id="8b210-117">**Найдите** файл манифеста надстройки и выберите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="8b210-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)

6. <span data-ttu-id="8b210-p103">Убедитесь, что надстройка установлена. Например, если надстройка вызывается командой, эта команда должна появиться на ленте или в контекстном меню. Если же у вас надстройка области задач, должна появиться область.</span><span class="sxs-lookup"><span data-stu-id="8b210-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.</span></span>
    

## <a name="sideload-an-office-add-in-in-office-online"></a><span data-ttu-id="8b210-122">Загрузка неопубликованной надстройки Office в Office Online</span><span class="sxs-lookup"><span data-stu-id="8b210-122">Sideload an Office Add-in on Office Online</span></span>


1. <span data-ttu-id="8b210-123">Откройте [Microsoft Office Online](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="8b210-123">Open [Microsoft Office Online](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="8b210-124">В разделе  **Начало работы с веб-приложениями** выберите **Excel**,  **Word** или **PowerPoint** и откройте новый документ.</span><span class="sxs-lookup"><span data-stu-id="8b210-124">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="8b210-125">Откройте вкладку  **Вставка** на ленте и в разделе **Надстройки** выберите **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="8b210-125">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="8b210-126">В диалоговом окне **Надстройки Office** откройте вкладку **МОИ НАДСТРОЙКИ** и выберите **Управление моими надстройками** > **Отправить надстройку**.</span><span class="sxs-lookup"><span data-stu-id="8b210-126">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="8b210-128">**Найдите** файл манифеста надстройки и выберите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="8b210-128">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)

6. <span data-ttu-id="8b210-p104">Убедитесь, что надстройка установлена. Например, если надстройка вызывается командой, эта команда должна появиться на ленте или в контекстном меню. Если же у вас надстройка области задач, должна появиться область.</span><span class="sxs-lookup"><span data-stu-id="8b210-p104">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="8b210-133">Чтобы протестировать надстройку Office в Edge, введите **about:flags** в панели поиска Edge, чтобы открыть раздел "Параметры разработчика".</span><span class="sxs-lookup"><span data-stu-id="8b210-133">To test your Office Add-in with Edge, enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="8b210-134">Установите флажок **Разрешить замыкание на себя для localhost** и перезапустите Edge.</span><span class="sxs-lookup"><span data-stu-id="8b210-134">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![Параметр "Разрешить замыкание на себя для localhost" в Edge с установленным флажком.](../images/allow-localhost-loopback.png)

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="8b210-136">Загрузка неопубликованной настройки при использовании Visual Studio</span><span class="sxs-lookup"><span data-stu-id="8b210-136">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="8b210-137">Если вы разрабатываете надстройки с помощью Visual Studio, процесс загрузки неопубликованной надстройки будет аналогичным.</span><span class="sxs-lookup"><span data-stu-id="8b210-137">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="8b210-138">Единственное различие состоит в том, что необходимо обновить значение элемента **SourceURL** в манифесте, чтобы включить в него полный URL-адрес расположения, в котором развернута надстройка.</span><span class="sxs-lookup"><span data-stu-id="8b210-138">If you're using Visual Studio to develop your add-in, the process to sideload is similar. The only difference is that you will have to update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="8b210-139">Хотя неопубликованные надстройки можно загружать из Visual Studio в Office Online, их невозможно отлаживать из Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="8b210-139">Although you can sideload add-ins from Visual Studio to Office Online, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="8b210-140">Для этого вам потребуются средства отладки браузера.</span><span class="sxs-lookup"><span data-stu-id="8b210-140">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="8b210-141">Дополнительные сведения см. в статье [Отладка надстроек в Office Online](debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="8b210-141">For more information, see [Debug add-ins in Office Online](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="8b210-142">В Visual Studio откройте окно **Свойства**, выбрав **Вид** -> **Окно свойств**.</span><span class="sxs-lookup"><span data-stu-id="8b210-142">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="8b210-143">В **обозревателе решений** выберите веб-проект.</span><span class="sxs-lookup"><span data-stu-id="8b210-143">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="8b210-144">В окне **Свойства** отобразятся свойства проекта.</span><span class="sxs-lookup"><span data-stu-id="8b210-144">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="8b210-145">В окне "Свойства" скопируйте **URL-адрес SSL**.</span><span class="sxs-lookup"><span data-stu-id="8b210-145">In the  Properties window, copy the value of the SSL URL property. An example ishttps://localhost:44300/.</span></span>
4. <span data-ttu-id="8b210-146">В проекте надстройки откройте XML-файл манифеста.</span><span class="sxs-lookup"><span data-stu-id="8b210-146">In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml”.</span></span> <span data-ttu-id="8b210-147">Убедитесь, что вы изменяете исходный XML-файл.</span><span class="sxs-lookup"><span data-stu-id="8b210-147">Be sure you are editing the source XML.</span></span> <span data-ttu-id="8b210-148">Для проектов некоторых типов в Visual Studio откроется визуальное представление XML-файла, которое не будет работать на следующем шаге.</span><span class="sxs-lookup"><span data-stu-id="8b210-148">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="8b210-149">Найдите и замените все экземпляры **~remoteAppUrl/** только что скопированным URL-адресом SSL.</span><span class="sxs-lookup"><span data-stu-id="8b210-149">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="8b210-150">В зависимости от типа проекта отобразится несколько вариантов замены, и появятся новые URL-адреса, похожие на `https://localhost:44300/Home.html`.</span><span class="sxs-lookup"><span data-stu-id="8b210-150">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="8b210-151">Сохраните XML-файл.</span><span class="sxs-lookup"><span data-stu-id="8b210-151">3-  Save the XML manifest file.</span></span>
7. <span data-ttu-id="8b210-152">Щелкните веб-проект правой кнопкой мыши и выберите **Отладка** -> **Запустить новый экземпляр**.</span><span class="sxs-lookup"><span data-stu-id="8b210-152">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="8b210-153">Веб-проект будет выполнен без запуска Office.</span><span class="sxs-lookup"><span data-stu-id="8b210-153">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="8b210-154">В Office Online загрузите неопубликованную надстройку согласно инструкциям, приведенным выше в разделе [Загрузка неопубликованной надстройки Office в Office Online](#sideload-an-office-add-in-in-office-online).</span><span class="sxs-lookup"><span data-stu-id="8b210-154">From Office Online, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office Online](#sideload-an-office-add-in-in-office-online).</span></span>
