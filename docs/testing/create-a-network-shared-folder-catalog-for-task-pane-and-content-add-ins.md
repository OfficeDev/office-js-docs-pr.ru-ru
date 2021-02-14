---
title: Загрузка неогрузки надстройки Office для тестирования из сетевой сети
description: Узнайте, как загрузка неогрузки надстройки Office для тестирования из сетевой сети
ms.date: 06/02/2020
localization_priority: Normal
ms.openlocfilehash: 7e584b5543d988ed51f932254d48981d51afa0fc
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237975"
---
# <a name="sideload-office-add-ins-for-testing-from-a-network-share"></a><span data-ttu-id="75f2f-103">Загрузка неогрузки надстройки Office для тестирования из сетевой сети</span><span class="sxs-lookup"><span data-stu-id="75f2f-103">Sideload Office Add-ins for testing from a network share</span></span>

<span data-ttu-id="75f2f-104">Вы можете протестировать надстройку Office в клиенте Office, который находится в Windows, опубликуя манифест в сетевой папке (инструкции ниже).</span><span class="sxs-lookup"><span data-stu-id="75f2f-104">You can test an Office Add-in in an Office client that is on Windows by publishing the manifest to a network file share (instructions below).</span></span> <span data-ttu-id="75f2f-105">Этот вариант развертывания предназначен для использования после завершения разработки и тестирования на localhost и для тестирования надстройки с нелокализованного сервера или облачной учетной записи.</span><span class="sxs-lookup"><span data-stu-id="75f2f-105">This deployment option is intended to be used when you have completed development and testing on a localhost and want to test the add-in from a non-local server or cloud account.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="75f2f-106">Развертывание по сетевой доской не поддерживается для производственных надстройок. Этот метод имеет следующие ограничения:</span><span class="sxs-lookup"><span data-stu-id="75f2f-106">Deployment by network share is not supported for production add-ins. This method has the following limitations:</span></span>
> 
> - <span data-ttu-id="75f2f-107">Надстройка может быть установлена только на компьютерах с Windows.</span><span class="sxs-lookup"><span data-stu-id="75f2f-107">The add-in can only be installed on Windows computers.</span></span>
> - <span data-ttu-id="75f2f-108">Если новая версия надстройки изменяет ленту, каждому пользователю придется переустановить надстройки.</span><span class="sxs-lookup"><span data-stu-id="75f2f-108">If a new version of an add-in changes the ribbon, each user will have to reinstall the add-in.</span></span>


> [!NOTE]
> <span data-ttu-id="75f2f-109">Если проект надстройки создан с помощью достаточно новой версии [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office), неопубликованная надстройка автоматически загружается в классический клиент Office, когда вы запускаете команду `npm start`.</span><span class="sxs-lookup"><span data-stu-id="75f2f-109">If your add-in project was created with a sufficiently recent version of the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), the add-in will automatically sideload in the Office desktop client when you run `npm start`.</span></span>

<span data-ttu-id="75f2f-110">Эта статья относится только к тестированию надстройки Word, Excel, PowerPoint и Project и только в Windows.</span><span class="sxs-lookup"><span data-stu-id="75f2f-110">This article applies only to testing Word, Excel, PowerPoint, and Project add-ins and only on Windows.</span></span> <span data-ttu-id="75f2f-111">Для выполнения тестирования на другой платформе или тестирования надстроек Outlook см. одну из указанных ниже тем о загрузке неопубликованных надстроек.</span><span class="sxs-lookup"><span data-stu-id="75f2f-111">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="75f2f-112">Загрузка неопубликованных надстроек Office в Office в Интернете для тестирования</span><span class="sxs-lookup"><span data-stu-id="75f2f-112">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="75f2f-113">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="75f2f-113">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="75f2f-114">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="75f2f-114">Sideload Outlook add-ins for testing</span></span>](../outlook/sideload-outlook-add-ins-for-testing.md)

<span data-ttu-id="75f2f-115">В приведенном ниже видео показано, как загрузить неопубликованную надстройку в классическое приложение Office или Office в Интернете с помощью каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="75f2f-115">The following video walks you through the process of sideloading your add-in in Office on the web or desktop using a shared folder catalog.</span></span>  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a><span data-ttu-id="75f2f-116">Общий доступ к папке</span><span class="sxs-lookup"><span data-stu-id="75f2f-116">Share a folder</span></span>

1. <span data-ttu-id="75f2f-117">На том компьютере с Windows, где должна размещаться надстройка, перейдите к родительской папке или диску с папкой, которую требуется использовать в качестве каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="75f2f-117">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="75f2f-118">Откройте контекстное меню для папки, которую нужно использовать в качестве каталога общих папок (щелкните папку правой кнопкой мыши), и выберите пункт **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-118">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="75f2f-119">В диалоговом окне **Свойства** откройте вкладку **Доступ** и нажмите кнопку **Общий доступ**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-119">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![диалоговое окно "Свойства" папки с выделенной вкладкой "Доступ" и кнопкой "Общий доступ"](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="75f2f-121">В диалоговом окне **Доступ к сети** добавьте себя и любых других пользователей и/или группы пользователей, которым следует предоставить доступ к вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="75f2f-121">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="75f2f-122">Вам потребуются разрешения на **чтение и запись** папки.</span><span class="sxs-lookup"><span data-stu-id="75f2f-122">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="75f2f-123">Завершив выбор пользователей, которым предоставляется совместный доступ, нажмите кнопку **Поделиться**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-123">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="75f2f-124">Когда появится подтверждение, что **папка открыта для общего доступа**, запишите полный сетевой путь, отображаемый сразу после имени папки.</span><span class="sxs-lookup"><span data-stu-id="75f2f-124">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="75f2f-125">(Это значение нужно вводить как **URL-адрес каталога** при [указании общей папки в качестве доверенного каталога](#specify-the-shared-folder-as-a-trusted-catalog), как описано в следующем разделе этой статьи). Нажмите кнопку **Готово**, чтобы закрыть диалоговое окно **Доступ к сети**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-125">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Диалоговое окно "Доступ к сети" с выделенным путем для общего доступа](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="75f2f-127">Нажмите кнопку **Закрыть**, чтобы закрыть диалоговое окно **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-127">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="75f2f-128">Указание общей папки в качестве доверенного каталога</span><span class="sxs-lookup"><span data-stu-id="75f2f-128">Specify the shared folder as a trusted catalog</span></span>

### <a name="configure-the-trust-manually"></a><span data-ttu-id="75f2f-129">Настройка доверия вручную</span><span class="sxs-lookup"><span data-stu-id="75f2f-129">Configure the trust manually</span></span>

1. <span data-ttu-id="75f2f-130">Откройте новый документ в Excel, Word, PowerPoint или Project.</span><span class="sxs-lookup"><span data-stu-id="75f2f-130">Open a new document in Excel, Word, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="75f2f-131">Перейдите на вкладку **Файл**, а затем выберите **Параметры**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-131">Choose the **File** tab, and then choose **Options**.</span></span>

3. <span data-ttu-id="75f2f-132">Выберите **Центр управления безопасностью**, а затем нажмите кнопку **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-132">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>

4. <span data-ttu-id="75f2f-133">Выберите пункт **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-133">Choose **Trusted Add-in Catalogs**.</span></span>

5. <span data-ttu-id="75f2f-134">В поле **URL-адрес каталога** введите полный сетевой путь к папке, к которой вы ранее предоставили [общий доступ](#share-a-folder).</span><span class="sxs-lookup"><span data-stu-id="75f2f-134">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="75f2f-135">Если вы не записали полный сетевой путь к папке при предоставлении к ней общего доступа, его можно получить в диалоговом окне **Свойства** папки, как показано на снимке экрана ниже.</span><span class="sxs-lookup"><span data-stu-id="75f2f-135">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span>

    ![диалоговое окно "Свойства" папки с выделенной вкладкой "Доступ" и сетевым путем](../images/sideload-windows-properties-dialog-2.png)

6. <span data-ttu-id="75f2f-137">После ввода полного сетевого пути доступа к папке в поле **URL-адрес каталога**, нажмите кнопку **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-137">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="75f2f-138">Установите флажок **Показывать в меню** для только что добавленного элемента и нажмите кнопку **ОК**, чтобы закрыть диалоговое окно **Центр управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-138">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Диалоговое окно "Центр управления безопасностью" с выбранным каталогом](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="75f2f-140">Чтобы закрыть **диалоговое** окно **"Параметры",** выберите кнопку "ОК".</span><span class="sxs-lookup"><span data-stu-id="75f2f-140">Choose the **OK** button to close the **Options** dialog window.</span></span>

9. <span data-ttu-id="75f2f-141">Закройте и снова откройте приложение Office, чтобы изменения вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="75f2f-141">Close and reopen the Office application so your changes will take effect.</span></span>

### <a name="configure-the-trust-with-a-registry-script"></a><span data-ttu-id="75f2f-142">Настройка доверия с помощью сценария реестра</span><span class="sxs-lookup"><span data-stu-id="75f2f-142">Configure the trust with a Registry script</span></span>

1. <span data-ttu-id="75f2f-143">В текстовом редакторе создайте файл с именем TrustNetworkShareCatalog.reg.</span><span class="sxs-lookup"><span data-stu-id="75f2f-143">In a text editor, create a file named TrustNetworkShareCatalog.reg.</span></span>

2. <span data-ttu-id="75f2f-144">Добавьте следующее содержимое в файл:</span><span class="sxs-lookup"><span data-stu-id="75f2f-144">Add the following content to the file:</span></span>

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```
3. <span data-ttu-id="75f2f-145">Используйте одно из многочисленных средств создания GUID в Интернете, например [Генератор GUID](https://guidgenerator.com/), для создания случайного GUID и в файле TrustNetworkShareCatalog.reg замените строку "-random-GUID-here-" *в обоих местах* идентификатором GUID.</span><span class="sxs-lookup"><span data-stu-id="75f2f-145">Use one of the many online GUID generation tools, such as [GUID Generator](https://guidgenerator.com/), to generate a random GUID, and within the TrustNetworkShareCatalog.reg file, replace the string "-random-GUID-here-" *in both places* with the GUID.</span></span> <span data-ttu-id="75f2f-146">(Символы `{}` должны сохраняться.)</span><span class="sxs-lookup"><span data-stu-id="75f2f-146">(The enclosing `{}` symbols should remain.)</span></span>

4. <span data-ttu-id="75f2f-147">Замените значение `Url` полным сетевым путем к папке, к которой вы ранее предоставили [общий доступ](#share-a-folder).</span><span class="sxs-lookup"><span data-stu-id="75f2f-147">Replace the `Url` value with the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="75f2f-148">(Обратите внимание, что все знаки `\` в URL-адресе должны дублироваться.) Если вы не записали полный сетевой путь к папке при предоставлении к ней общего доступа, его можно получить в диалоговом окне **Свойства** папки, как показано на снимке экрана ниже.</span><span class="sxs-lookup"><span data-stu-id="75f2f-148">(Note that any `\` characters in the URL must be doubled.) If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span>

    ![диалоговое окно "Свойства" папки с выделенной вкладкой "Доступ" и сетевым путем](../images/sideload-windows-properties-dialog-2.png)

5. <span data-ttu-id="75f2f-150">Файл теперь должен выглядеть так, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="75f2f-150">The file should now look like the following.</span></span> <span data-ttu-id="75f2f-151">Сохраните его.</span><span class="sxs-lookup"><span data-stu-id="75f2f-151">Save it.</span></span>

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

6. <span data-ttu-id="75f2f-152">Закройте *все* приложения Office.</span><span class="sxs-lookup"><span data-stu-id="75f2f-152">Close *all* Office applications.</span></span>

7. <span data-ttu-id="75f2f-153">Запустите файл TrustNetworkShareCatalog.reg как любой исполняемый файл, например, дважды щелкнув его.</span><span class="sxs-lookup"><span data-stu-id="75f2f-153">Run the TrustNetworkShareCatalog.reg just as you would any executable, such as double-clicking it.</span></span>

## <a name="sideload-your-add-in"></a><span data-ttu-id="75f2f-154">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="75f2f-154">Sideload your add-in</span></span>

1. <span data-ttu-id="75f2f-155">XML-файл манифеста тестируемой надстройки необходимо поместить в каталог общих папок.</span><span class="sxs-lookup"><span data-stu-id="75f2f-155">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="75f2f-156">Обратите внимание, что вы развертываете веб-приложение непосредственно на веб-сервере.</span><span class="sxs-lookup"><span data-stu-id="75f2f-156">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="75f2f-157">Не забудьте указать URL-адрес в элементе **SourceLocation** файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="75f2f-157">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="75f2f-158">В Excel, Word или PowerPoint откройте на ленте вкладку **Вставка** и выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-158">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> <span data-ttu-id="75f2f-159">В Project выберите **Мои надстройки** на вкладке **Project** ленты.</span><span class="sxs-lookup"><span data-stu-id="75f2f-159">In Project, select **My Add-ins** on the **Project** tab of the ribbon.</span></span>

3. <span data-ttu-id="75f2f-160">Нажмите **ОБЩАЯ ПАПКА** в верхней части диалогового окна **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="75f2f-160">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="75f2f-161">Выберите имя надстройки и нажмите кнопку **Добавить**, чтобы вставить надстройку.</span><span class="sxs-lookup"><span data-stu-id="75f2f-161">Select the name of the add-in and choose **Add** to insert the add-in.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="75f2f-162">Удаление неогруженной надстройки</span><span class="sxs-lookup"><span data-stu-id="75f2f-162">Remove a sideloaded add-in</span></span>

<span data-ttu-id="75f2f-163">Вы можете удалить ранее загруженную неогруженную надстройку, с помощью очистки кэша Office на компьютере.</span><span class="sxs-lookup"><span data-stu-id="75f2f-163">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="75f2f-164">Сведения о том, как очистить кэш в Windows, можно найти в статье ["Очистка кэша Office".](clear-cache.md#clear-the-office-cache-on-windows)</span><span class="sxs-lookup"><span data-stu-id="75f2f-164">Details on how to clear the cache on Windows can be found in the article [Clear the Office cache](clear-cache.md#clear-the-office-cache-on-windows).</span></span>

## <a name="see-also"></a><span data-ttu-id="75f2f-165">См. также</span><span class="sxs-lookup"><span data-stu-id="75f2f-165">See also</span></span>

- [<span data-ttu-id="75f2f-166">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="75f2f-166">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="75f2f-167">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="75f2f-167">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="75f2f-168">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="75f2f-168">Publish your Office Add-in</span></span>](../publish/publish.md)
