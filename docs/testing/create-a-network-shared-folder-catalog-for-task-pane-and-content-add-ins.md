---
title: Загрузка неопубликованных надстроек Office для тестирования
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 6ee8e4e9a2413b34cb8991b09d61e16888a0e6a6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640024"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="afdaa-102">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="afdaa-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="afdaa-103">Вы можете установить надстройку Office для тестирования в работающем под Windows клиенте Office, опубликовав манифест в сетевом файловом ресурсе (см. инструкции ниже).</span><span class="sxs-lookup"><span data-stu-id="afdaa-103">You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share.</span></span>

> [!NOTE]
> <span data-ttu-id="afdaa-p101">Если для создания проекта надстройки использовалось средство [**yo office** ](https://github.com/OfficeDev/generator-office), то существует альтернативный способ загрузки неопубликованных надстроек, который может оказаться оптимальным в вашем случае. Для получения дополнительной информации см. [Загрузка неопубликованных надстроек Office с помощью команды sideload](sideload-office-addin-using-sideload-command.md).</span><span class="sxs-lookup"><span data-stu-id="afdaa-p101">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you. For details, see [Sideload Office Add-ins using the sideload command](sideload-office-addin-using-sideload-command.md).</span></span>

<span data-ttu-id="afdaa-p102">Эта статья относится только к тестированию надстроек Word, Excel или PowerPoint в Windows. Для выполнения тестирования на другой платформе или тестирования надстроек Outlook, см. одну из следующих тем, касающихся загрузки неопубликованных надстроек:</span><span class="sxs-lookup"><span data-stu-id="afdaa-p102">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows. If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="afdaa-108">Загрузка неопубликованных надстроек Office в Office Online для тестирования</span><span class="sxs-lookup"><span data-stu-id="afdaa-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="afdaa-109">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="afdaa-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="afdaa-110">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="afdaa-110">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="afdaa-111">В следующем видео вы ознакомитесь с процессом загрузки вашей неопубликованной надстройки для классической версии Office или Office Online с помощью каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="afdaa-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="afdaa-112">Совместное использование папки</span><span class="sxs-lookup"><span data-stu-id="afdaa-112">Share a folder</span></span>

1. <span data-ttu-id="afdaa-113">Воспользовавшись проводником на компьютере с установленной ОС Windows, где следует разместить надстройку, перейдите к родительской папке или диску с папкой, которую требуется использовать в качестве каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="afdaa-113">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="afdaa-114">Откройте контекстное меню для папки, которую предполагается использовать в качестве каталога общей папки (щелкните на папке правой кнопкой мыши), и выберите опцию **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-114">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="afdaa-115">В диалоговом окне **Свойства** откройте вкладку **Общий доступ**, а затем нажмите кнопку **Предоставить общий доступ**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-115">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![Диалоговое окно «Свойства» с вкладкой «Общий доступ» и выделенной кнопкой «Предоставить общий доступ»](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="afdaa-117">В диалоговом окне **Сетевой доступ** добавьте себя и любых других пользователей и/или группы пользователей, которым следует предоставить доступ к вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="afdaa-117">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="afdaa-118">Вам потребуется разрешение на **чтение и запись** папки.</span><span class="sxs-lookup"><span data-stu-id="afdaa-118">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="afdaa-119">Завершив выбор пользователей, которым предоставляется доступ к папке, нажмите на кнопку **Предоставить общий доступ** .</span><span class="sxs-lookup"><span data-stu-id="afdaa-119">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="afdaa-120">При появлении подтверждения **Общий доступ к папке предоставлен**, запишите полный сетевой путь, который отображается сразу после имени папки.</span><span class="sxs-lookup"><span data-stu-id="afdaa-120">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="afdaa-121">(Это значение следует вводить, как **URL-адрес каталога** при [указании общей папки в качестве доверенного каталога](#specify-the-shared-folder-as-a-trusted-catalog), как описано в следующем разделе этой статьи). Нажмите на кнопку **Готово**, чтобы закрыть диалоговое окно **Сетевой доступ**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-121">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Диалоговое окно «Сетевой доступ» с выделенным путем доступа к общей папке](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="afdaa-123">Нажмите на кнопку **Закрыть**, чтобы закрыть диалоговое окно **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-123">Choose the **Close** button to close the **Workbook Connections** dialog box.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="afdaa-124">Укажите общую папку в качестве каталога общих папок</span><span class="sxs-lookup"><span data-stu-id="afdaa-124">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="afdaa-125">Откройте новый документ в Excel, Word или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="afdaa-125">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="afdaa-126">Перейдите на вкладку **Файл**, а затем выберите **Параметры**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="afdaa-127">Выберите опцию **Центр управления безопасностью**, а затем нажмите на кнопку **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="afdaa-128">Выберите пункт **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-128">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="afdaa-129">В поле **URL-адрес каталога** введите полный сетевой путь доступа к папке, [общий доступ](#share-a-folder) к которой был предоставлен ранее.</span><span class="sxs-lookup"><span data-stu-id="afdaa-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="afdaa-130">Если при предоставлении общего доступа к папке ее полный сетевой путь зарегистрировать не удалось, то узнать его можно в диалоговом окне **Свойства** папки, как показано на следующем снимке экрана.</span><span class="sxs-lookup"><span data-stu-id="afdaa-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![Диалоговое окно «Свойства» папки с вкладкой «Общий доступ» и выденным сетевым путем доступа](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="afdaa-132">После ввода полного сетевого пути доступа к папке в поле **URL-адрес каталога** нажмите на кнопку **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="afdaa-133">Установите флажок **Показывать в меню** для недавно добавленного элемента, а затем нажмите на кнопку **ОК**, чтобы закрыть диалоговое окно **Центр управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Диалоговое окно «Центр управления безопасностью» с выбранным каталогом](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="afdaa-135">Нажмите на кнопку **ОК**, чтобы закрыть диалоговое окно **Параметры Word**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-135">Choose the  **OK** button to close the **Internet Options** dialog box.</span></span>

9. <span data-ttu-id="afdaa-136">Чтобы изменения вступили в силу, закройте и снова откройте приложение Office.</span><span class="sxs-lookup"><span data-stu-id="afdaa-136">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="afdaa-137">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="afdaa-137">Sideload your add-in</span></span>


1. <span data-ttu-id="afdaa-138">Поместите XML-файл манифеста любой надстройки, которую вы тестируете, в каталоге общих папок.</span><span class="sxs-lookup"><span data-stu-id="afdaa-138">Put the manifest file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="afdaa-139">Обратите внимание на то, вы развертываете само веб-приложение на веб-сервере.</span><span class="sxs-lookup"><span data-stu-id="afdaa-139">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="afdaa-140">Укажите URL-адрес в элементе **SourceLocation** файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="afdaa-140">Deploy the web application itself to a web server and specify the URL in the  **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="afdaa-141">В Excel, Word или PowerPoint откройте на ленте вкладку **Вставка** и выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-141">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="afdaa-142">Выберите опцию **ОБЩАЯ ПАПКА** в верхней части диалогового окна **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="afdaa-142">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="afdaa-143">Выберите имя надстройки и нажмите на кнопку **ОК**, чтобы вставить надстройку.</span><span class="sxs-lookup"><span data-stu-id="afdaa-143">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="afdaa-144">См. также</span><span class="sxs-lookup"><span data-stu-id="afdaa-144">See also</span></span>

- [<span data-ttu-id="afdaa-145">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="afdaa-145">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="afdaa-146">Публикация надстроек Office</span><span class="sxs-lookup"><span data-stu-id="afdaa-146">Publish your Office Add-in</span></span>](../publish/publish.md)
    
