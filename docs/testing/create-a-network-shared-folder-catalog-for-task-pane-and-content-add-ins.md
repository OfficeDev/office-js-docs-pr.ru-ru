---
title: Загрузка неопубликованных надстроек Office для тестирования
description: ''
ms.date: 12/06/2019
localization_priority: Priority
ms.openlocfilehash: bb926b09d9381574d22e7634a578adac141e1f8f
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814481"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="d9798-102">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="d9798-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="d9798-103">Вы можете установить надстройку Office для тестирования в клиенте Office, запущенном в Windows, используя каталог общих папок для публикации манифеста в сетевом файловом ресурсе.</span><span class="sxs-lookup"><span data-stu-id="d9798-103">You can install an Office Add-in for testing in an Office client running on Windows by publishing the manifest to a network file share (instructions below).</span></span>

> [!NOTE]
> <span data-ttu-id="d9798-104">Если проект надстройки создан с помощью достаточно новой версии [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office), неопубликованная надстройка автоматически загружается в классический клиент Office, когда вы запускаете команду `npm start`.</span><span class="sxs-lookup"><span data-stu-id="d9798-104">If your add-in project was created with a sufficiently recent version of the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), the add-in will automatically sideload in the Office desktop client when you run `npm start`.</span></span>

<span data-ttu-id="d9798-105">Эта статья относится только к тестированию надстроек Word, Excel, PowerPoint и Project в Windows.</span><span class="sxs-lookup"><span data-stu-id="d9798-105">This article applies only to testing Word, Excel, PowerPoint, and Project add-ins on Windows.</span></span> <span data-ttu-id="d9798-106">Для выполнения тестирования на другой платформе или тестирования надстроек Outlook см. одну из указанных ниже тем о загрузке неопубликованных надстроек.</span><span class="sxs-lookup"><span data-stu-id="d9798-106">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="d9798-107">Загрузка неопубликованных надстроек Office в Office в Интернете для тестирования</span><span class="sxs-lookup"><span data-stu-id="d9798-107">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="d9798-108">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="d9798-108">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="d9798-109">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="d9798-109">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

<span data-ttu-id="d9798-110">В приведенном ниже видео показано, как загрузить неопубликованную надстройку в классическое приложение Office или Office в Интернете с помощью каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="d9798-110">The following video walks you through the process of sideloading your add-in in Office on the web or desktop using a shared folder catalog.</span></span>  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a><span data-ttu-id="d9798-111">Общий доступ к папке</span><span class="sxs-lookup"><span data-stu-id="d9798-111">Share a folder</span></span>

1. <span data-ttu-id="d9798-112">На том компьютере с Windows, где должна размещаться надстройка, перейдите к родительской папке или диску с папкой, которую требуется использовать в качестве каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="d9798-112">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="d9798-113">Откройте контекстное меню для папки, которую нужно использовать в качестве каталога общих папок (щелкните папку правой кнопкой мыши), и выберите пункт **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="d9798-113">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="d9798-114">В диалоговом окне **Свойства** откройте вкладку **Доступ** и нажмите кнопку **Общий доступ**.</span><span class="sxs-lookup"><span data-stu-id="d9798-114">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![диалоговое окно "Свойства" папки с выделенной вкладкой "Доступ" и кнопкой "Общий доступ"](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="d9798-116">В диалоговом окне **Доступ к сети** добавьте себя и любых других пользователей и/или группы пользователей, которым следует предоставить доступ к вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="d9798-116">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="d9798-117">Вам потребуются разрешения на **чтение и запись** папки.</span><span class="sxs-lookup"><span data-stu-id="d9798-117">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="d9798-118">Завершив выбор пользователей, которым предоставляется совместный доступ, нажмите кнопку **Поделиться**.</span><span class="sxs-lookup"><span data-stu-id="d9798-118">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="d9798-119">Когда появится подтверждение, что **папка открыта для общего доступа**, запишите полный сетевой путь, отображаемый сразу после имени папки.</span><span class="sxs-lookup"><span data-stu-id="d9798-119">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="d9798-120">(Это значение нужно вводить как **URL-адрес каталога** при [указании общей папки в качестве доверенного каталога](#specify-the-shared-folder-as-a-trusted-catalog), как описано в следующем разделе этой статьи). Нажмите кнопку **Готово**, чтобы закрыть диалоговое окно **Доступ к сети**.</span><span class="sxs-lookup"><span data-stu-id="d9798-120">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Диалоговое окно "Доступ к сети" с выделенным путем для общего доступа](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="d9798-122">Нажмите кнопку **Закрыть**, чтобы закрыть диалоговое окно **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="d9798-122">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="d9798-123">Указание общей папки в качестве доверенного каталога</span><span class="sxs-lookup"><span data-stu-id="d9798-123">Specify the shared folder as a trusted catalog</span></span> 

### <a name="configure-the-trust-manually"></a><span data-ttu-id="d9798-124">Настройка доверия вручную</span><span class="sxs-lookup"><span data-stu-id="d9798-124">Configure the trust manually</span></span>
      
1. <span data-ttu-id="d9798-125">Откройте новый документ в Excel, Word, PowerPoint или Project.</span><span class="sxs-lookup"><span data-stu-id="d9798-125">Open a new document in Excel, Word, PowerPoint, or Project.</span></span>
    
2. <span data-ttu-id="d9798-126">Перейдите на вкладку **Файл**, а затем выберите **Параметры**.</span><span class="sxs-lookup"><span data-stu-id="d9798-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="d9798-127">Выберите **Центр управления безопасностью**, а затем нажмите кнопку **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="d9798-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="d9798-128">Выберите пункт **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="d9798-128">Choose **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="d9798-129">В поле **URL-адрес каталога** введите полный сетевой путь к папке, к которой вы ранее предоставили [общий доступ](#share-a-folder).</span><span class="sxs-lookup"><span data-stu-id="d9798-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="d9798-130">Если вы не записали полный сетевой путь к папке при предоставлении к ней общего доступа, его можно получить в диалоговом окне **Свойства** папки, как показано на снимке экрана ниже.</span><span class="sxs-lookup"><span data-stu-id="d9798-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![диалоговое окно "Свойства" папки с выделенной вкладкой "Доступ" и сетевым путем](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="d9798-132">После ввода полного сетевого пути доступа к папке в поле **URL-адрес каталога**, нажмите кнопку **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="d9798-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="d9798-133">Установите флажок **Показывать в меню** для только что добавленного элемента и нажмите кнопку **ОК**, чтобы закрыть диалоговое окно **Центр управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="d9798-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Диалоговое окно "Центр управления безопасностью" с выбранным каталогом](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="d9798-135">Нажмите кнопку **ОК**, чтобы закрыть диалоговое окно **Параметры Word**.</span><span class="sxs-lookup"><span data-stu-id="d9798-135">Choose the **OK** button to close the **Word Options** dialog window.</span></span>

9. <span data-ttu-id="d9798-136">Закройте и снова откройте приложение Office, чтобы изменения вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="d9798-136">Close and reopen the Office application so your changes will take effect.</span></span>

### <a name="configure-the-trust-with-a-registry-script"></a><span data-ttu-id="d9798-137">Настройка доверия с помощью сценария реестра</span><span class="sxs-lookup"><span data-stu-id="d9798-137">Configure the trust with a Registry script</span></span>

1. <span data-ttu-id="d9798-138">В текстовом редакторе создайте файл с именем TrustNetworkShareCatalog.reg.</span><span class="sxs-lookup"><span data-stu-id="d9798-138">In a text editor, such as Notepad, create a file named ItemMetadata.xml.</span></span> 

2. <span data-ttu-id="d9798-139">Добавьте следующее содержимое в файл:</span><span class="sxs-lookup"><span data-stu-id="d9798-139">Add the following content to the file:</span></span>

    ```
    Windows Registry Editor Version 5.00
    
    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```
3. <span data-ttu-id="d9798-140">Используйте одно из многочисленных средств создания GUID в Интернете, например [Генератор GUID](https://guidgenerator.com/), для создания случайного GUID и в файле TrustNetworkShareCatalog.reg замените строку "-random-GUID-here-" *в обоих местах* идентификатором GUID.</span><span class="sxs-lookup"><span data-stu-id="d9798-140">Use one of the many online GUID generation tools, such as [GUID Generator](https://guidgenerator.com/), to generate a random GUID, and within the TrustNetworkShareCatalog.reg file, replace the string "-random-GUID-here-" *in both places* with the GUID.</span></span> <span data-ttu-id="d9798-141">(Символы `{}` должны сохраняться.)</span><span class="sxs-lookup"><span data-stu-id="d9798-141">(The enclosing `{}` symbols should remain.)</span></span>

4. <span data-ttu-id="d9798-142">Замените значение `Url` полным сетевым путем к папке, к которой вы ранее предоставили [общий доступ](#share-a-folder).</span><span class="sxs-lookup"><span data-stu-id="d9798-142">In the `Url` box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="d9798-143">(Обратите внимание, что все знаки `\` в URL-адресе должны дублироваться.) Если вы не записали полный сетевой путь к папке при предоставлении к ней общего доступа, его можно получить в диалоговом окне **Свойства** папки, как показано на снимке экрана ниже.</span><span class="sxs-lookup"><span data-stu-id="d9798-143">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![диалоговое окно "Свойства" папки с выделенной вкладкой "Доступ" и сетевым путем](../images/sideload-windows-properties-dialog-2.png)
    
5. <span data-ttu-id="d9798-145">Файл теперь должен выглядеть так, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="d9798-145">The method should now look like the following.</span></span> <span data-ttu-id="d9798-146">Сохраните его.</span><span class="sxs-lookup"><span data-stu-id="d9798-146">Save it.</span></span>

    ```
    Windows Registry Editor Version 5.00
    
    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

6. <span data-ttu-id="d9798-147">Закройте *все* приложения Office.</span><span class="sxs-lookup"><span data-stu-id="d9798-147">Close all Office applications, including Excel.</span></span>

7. <span data-ttu-id="d9798-148">Запустите файл TrustNetworkShareCatalog.reg как любой исполняемый файл, например, дважды щелкнув его.</span><span class="sxs-lookup"><span data-stu-id="d9798-148">Run the TrustNetworkShareCatalog.reg just as you would any executable, such as double-clicking it.</span></span>

## <a name="sideload-your-add-in"></a><span data-ttu-id="d9798-149">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="d9798-149">Sideload your add-in</span></span>

1. <span data-ttu-id="d9798-150">XML-файл манифеста тестируемой надстройки необходимо поместить в каталог общих папок.</span><span class="sxs-lookup"><span data-stu-id="d9798-150">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="d9798-151">Обратите внимание, что вы развертываете веб-приложение непосредственно на веб-сервере.</span><span class="sxs-lookup"><span data-stu-id="d9798-151">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="d9798-152">Не забудьте указать URL-адрес в элементе **SourceLocation** файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="d9798-152">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="d9798-153">В Excel, Word или PowerPoint откройте на ленте вкладку **Вставка** и выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="d9798-153">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> <span data-ttu-id="d9798-154">В Project выберите **Мои надстройки** на вкладке **Project** ленты.</span><span class="sxs-lookup"><span data-stu-id="d9798-154">In Project, select **My Add-ins** on the **Project** tab of the ribbon.</span></span> 

3. <span data-ttu-id="d9798-155">Нажмите **ОБЩАЯ ПАПКА** в верхней части диалогового окна **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="d9798-155">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="d9798-156">Выберите имя надстройки и нажмите кнопку **Добавить**, чтобы вставить надстройку.</span><span class="sxs-lookup"><span data-stu-id="d9798-156">Select the name of the add-in and choose **Add** to insert the add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="d9798-157">См. также</span><span class="sxs-lookup"><span data-stu-id="d9798-157">See also</span></span>

- [<span data-ttu-id="d9798-158">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="d9798-158">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="d9798-159">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="d9798-159">Publish your Office Add-in</span></span>](../publish/publish.md)
    
