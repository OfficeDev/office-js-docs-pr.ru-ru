---
title: Публикация надстроек области задач и контентных надстроек в каталоге приложений SharePoint
description: Чтобы предоставить доступ к надстройкам Office пользователям в организации, администраторы могут отправлять файлы манифестов надстроек Office в соответствующий каталог приложений.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 106dfd2b1610be92f1b53dc1644ff3f8c60c0543
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617032"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a><span data-ttu-id="046c1-103">Публикация надстроек области задач и контентных надстроек в каталоге приложений SharePoint</span><span class="sxs-lookup"><span data-stu-id="046c1-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="046c1-p101">Каталог приложений — это отдельное семейство веб-сайтов в веб-приложении SharePoint или клиенте SharePoint Online, в котором размещены библиотеки документов для надстроек Office и SharePoint. Администраторы могут отправлять в него файлы манифестов надстроек Office, чтобы пользователи в организации могли получить доступ к этим надстройкам. Когда администратор зарегистрирует каталог приложений как доверенный, пользователи смогут вставлять надстройки в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="046c1-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="046c1-106">Каталоги приложений в SharePoint не поддерживают функции надстроек, реализованные в узле `VersionOverrides` [манифеста надстройки](../develop/add-in-manifests.md), такие как команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="046c1-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="046c1-107">Чтобы публиковать надстройки для облачной или гибридной среды, рекомендуем использовать [централизованное развертывание через Центр администрирования Office 365](../publish/centralized-deployment.md).</span><span class="sxs-lookup"><span data-stu-id="046c1-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="046c1-108">Каталоги приложений SharePoint не поддерживаются в Office для Mac.</span><span class="sxs-lookup"><span data-stu-id="046c1-108">App catalogs on SharePoint are not supported in Office on Mac.</span></span> <span data-ttu-id="046c1-109">Для развертывания надстроек Office на клиентах Mac необходимо отправить их в [AppSource](/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="046c1-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="create-an-app-catalog"></a><span data-ttu-id="046c1-110">Создание каталога приложений</span><span class="sxs-lookup"><span data-stu-id="046c1-110">Create app catalog site</span></span>

<span data-ttu-id="046c1-111">Выполните действия, описанные в одном из указанных ниже разделов, чтобы создать каталог приложений в локальном сервере SharePoint Server или Office 365.</span><span class="sxs-lookup"><span data-stu-id="046c1-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a><span data-ttu-id="046c1-112">Создание каталога приложений в локальном сервере SharePoint Server</span><span class="sxs-lookup"><span data-stu-id="046c1-112">To create an app catalog for on-premises SharePoint Server</span></span>

<span data-ttu-id="046c1-113">Чтобы создать каталог приложений SharePoint, следуйте инструкциям в статье [Настройка сайта каталога приложений для веб-приложения](/sharepoint/administration/manage-the-app-catalog).</span><span class="sxs-lookup"><span data-stu-id="046c1-113">To create the SharePoint app catalog, follow the instructions at [Configure the App Catalog site for a web application](/sharepoint/administration/manage-the-app-catalog).</span></span>

<span data-ttu-id="046c1-114">После создания каталога приложений выполните инструкции [по публикации надстройки Office](#publish-an-office-add-in).</span><span class="sxs-lookup"><span data-stu-id="046c1-114">Once you have created the app catalog follow the steps to [publish an Office Add-in](#publish-an-office-add-in).</span></span>

### <a name="to-create-an-app-catalog-on-office-365"></a><span data-ttu-id="046c1-115">Создание каталога приложений в Office 365</span><span class="sxs-lookup"><span data-stu-id="046c1-115">To create an app catalog on Office 365</span></span>

1. <span data-ttu-id="046c1-116">Перейдите в Центр администрирования Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="046c1-116">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="046c1-117">Сведения о том, как найти Центр администрирования, см. в статье [Сведения о Центре администрирования Microsoft 365](/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="046c1-117">For information on how to find the admin center, see [About the Microsoft 365 admin center](/office365/admin/admin-overview/about-the-admin-center).</span></span>

2. <span data-ttu-id="046c1-118">На странице Центра администрирования Microsoft 365 разверните список **центров администрирования** и выберите пункт **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="046c1-118">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="046c1-119">Для создания каталога нужно использовать классический Центр администрирования SharePoint.</span><span class="sxs-lookup"><span data-stu-id="046c1-119">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="046c1-120">Если вы находитесь в новом Центре администрирования SharePoint, выберите пункт **Классический Центр администрирования SharePoint** в области слева.</span><span class="sxs-lookup"><span data-stu-id="046c1-120">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>

3. <span data-ttu-id="046c1-121">В области задач слева выберите пункт **приложения**.</span><span class="sxs-lookup"><span data-stu-id="046c1-121">In the left task pane, choose  **Apps**.</span></span>

4. <span data-ttu-id="046c1-122">На странице **приложения** выберите пункт **Каталог приложений**.</span><span class="sxs-lookup"><span data-stu-id="046c1-122">On the **apps** page, select **App Catalog**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="046c1-123">Если каталог приложений уже создан и отображается на этой странице, вы можете пропустить остальные этапы и перейти к следующему разделу этой статьи, чтобы опубликовать надстройку в каталоге.</span><span class="sxs-lookup"><span data-stu-id="046c1-123">If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.</span></span>

5. <span data-ttu-id="046c1-124">На странице **Сайт каталога приложений** нажмите кнопку **ОК**, чтобы принять параметр по умолчанию и создать сайт каталога приложений.</span><span class="sxs-lookup"><span data-stu-id="046c1-124">On the **App Catalog Site** page, select **OK** to accept the default option and create a new app catalog site.</span></span>

6. <span data-ttu-id="046c1-125">На странице **Создание семейства веб-сайтов каталога приложений** укажите название сайта каталога приложений.</span><span class="sxs-lookup"><span data-stu-id="046c1-125">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

7. <span data-ttu-id="046c1-126">Укажите **адрес веб-сайта**.</span><span class="sxs-lookup"><span data-stu-id="046c1-126">Specify the web site address.</span></span>

8. <span data-ttu-id="046c1-127">Укажите **администратора**.</span><span class="sxs-lookup"><span data-stu-id="046c1-127">Specify an **Administrator**.</span></span>

9. <span data-ttu-id="046c1-128">Для параметра **Квота ресурсов сервера** установите значение 0 (ноль).</span><span class="sxs-lookup"><span data-stu-id="046c1-128">Set the **Server Resource Quota** to 0 (zero).</span></span> <span data-ttu-id="046c1-129">(Квота ресурсов сервера связана с регулированием изолированных решений с низкой производительностью, но вы не будете устанавливать изолированные решения на сайте каталога приложений.)</span><span class="sxs-lookup"><span data-stu-id="046c1-129">(The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

10. <span data-ttu-id="046c1-130">Нажмите кнопку **OK**.</span><span class="sxs-lookup"><span data-stu-id="046c1-130">Choose **OK**.</span></span>

## <a name="publish-an-office-add-in"></a><span data-ttu-id="046c1-131">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="046c1-131">Publish an Office Add-in</span></span>

<span data-ttu-id="046c1-132">Выполните действия, описанные в одном из указанных ниже разделов, чтобы опубликовать надстройку Office в каталоге приложений в Office 365 или локальном сервере SharePoint Server.</span><span class="sxs-lookup"><span data-stu-id="046c1-132">Complete the steps in one of the following sections to publish an Office Add-in to an app catalog on Office 365 or on-premises SharePoint Server.</span></span>

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a><span data-ttu-id="046c1-133">Публикация надстройки Office в каталоге приложений SharePoint в Office 365</span><span class="sxs-lookup"><span data-stu-id="046c1-133">To publish an Office add-in to a SharePoint app catalog on Office 365</span></span>

1. <span data-ttu-id="046c1-134">Перейдите в Центр администрирования Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="046c1-134">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="046c1-135">Сведения о том, как найти Центр администрирования, см. в статье [Сведения о Центре администрирования Microsoft 365](/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="046c1-135">For information on how to find the admin center, see [About the Microsoft 365 admin center](/office365/admin/admin-overview/about-the-admin-center).</span></span>
2. <span data-ttu-id="046c1-136">На странице Центра администрирования Microsoft 365 разверните список **центров администрирования** и выберите пункт **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="046c1-136">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="046c1-137">Для создания каталога нужно использовать классический Центр администрирования SharePoint.</span><span class="sxs-lookup"><span data-stu-id="046c1-137">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="046c1-138">Если вы находитесь в новом Центре администрирования SharePoint, выберите пункт **Классический Центр администрирования SharePoint** в области слева.</span><span class="sxs-lookup"><span data-stu-id="046c1-138">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>
3. <span data-ttu-id="046c1-139">В области задач слева выберите пункт **приложения**.</span><span class="sxs-lookup"><span data-stu-id="046c1-139">In the left task pane, choose  **Apps**.</span></span>
4. <span data-ttu-id="046c1-140">На странице **приложения** выберите пункт **Каталог приложений**.</span><span class="sxs-lookup"><span data-stu-id="046c1-140">On the **apps** page, select **App Catalog**.</span></span>
5. <span data-ttu-id="046c1-141">Выберите элемент **Распределить приложения для Office**.</span><span class="sxs-lookup"><span data-stu-id="046c1-141">Choose **Distribute apps for Office**.</span></span>
6. <span data-ttu-id="046c1-142">На странице **Приложения для Office** выберите команду **Создать**.</span><span class="sxs-lookup"><span data-stu-id="046c1-142">In the **Apps for Office** page, choose **New**.</span></span>
7. <span data-ttu-id="046c1-143">В диалоговом окне **Добавление документа** нажмите кнопку **Выбрать файлы**.</span><span class="sxs-lookup"><span data-stu-id="046c1-143">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
8. <span data-ttu-id="046c1-144">Найдите и укажите файл [манифеста](../develop/add-in-manifests.md) для добавления и нажмите кнопку **Открыть**.</span><span class="sxs-lookup"><span data-stu-id="046c1-144">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
9. <span data-ttu-id="046c1-145">В диалоговом окне **Добавление документа** нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="046c1-145">In the **Add a document** dialog box, choose **OK**.</span></span>

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a><span data-ttu-id="046c1-146">Публикация надстройки в каталоге приложений с помощью локального сервера SharePoint Server</span><span class="sxs-lookup"><span data-stu-id="046c1-146">To publish an add-in to an app catalog with on-premises SharePoint Server</span></span>

1. <span data-ttu-id="046c1-147">Откройте страницу **Центр администрирования**.</span><span class="sxs-lookup"><span data-stu-id="046c1-147">Open the SharePoint Central Administration main page.</span></span>
2. <span data-ttu-id="046c1-148">В области задач слева выберите пункт **Приложения**.</span><span class="sxs-lookup"><span data-stu-id="046c1-148">In the left task pane, choose  **Apps**.</span></span>
3. <span data-ttu-id="046c1-149">На странице **Приложения** в разделе **Управление приложениями** выберите пункт **Управление каталогом приложений**.</span><span class="sxs-lookup"><span data-stu-id="046c1-149">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>
4. <span data-ttu-id="046c1-150">На странице **Управление каталогом приложений** убедитесь, что в поле выбора **Веб-приложение** выбрано правильное веб-приложение.</span><span class="sxs-lookup"><span data-stu-id="046c1-150">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>
5. <span data-ttu-id="046c1-151">Выберите URL-адрес в разделе **URL-адрес сайта**, чтобы открыть сайт каталога приложений.</span><span class="sxs-lookup"><span data-stu-id="046c1-151">Choose the URL under the **Site URL** to open the app catalog site.</span></span>
6. <span data-ttu-id="046c1-152">Выберите элемент **Распределить приложения для Office**.</span><span class="sxs-lookup"><span data-stu-id="046c1-152">Choose **Distribute apps for Office**.</span></span>
7. <span data-ttu-id="046c1-153">На странице **Приложения для Office** выберите команду **Создать**.</span><span class="sxs-lookup"><span data-stu-id="046c1-153">In the **Apps for Office** page, choose **New**.</span></span>
8. <span data-ttu-id="046c1-154">В диалоговом окне **Добавление документа** нажмите кнопку **Выбрать файлы**.</span><span class="sxs-lookup"><span data-stu-id="046c1-154">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
9. <span data-ttu-id="046c1-155">Найдите и укажите файл [манифеста](../develop/add-in-manifests.md) для добавления и нажмите кнопку **Открыть**.</span><span class="sxs-lookup"><span data-stu-id="046c1-155">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
10. <span data-ttu-id="046c1-156">В диалоговом окне **Добавление документа** нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="046c1-156">In the **Add a document** dialog box, choose **OK**.</span></span>

## <a name="insert-office-add-ins-from-the-app-catalog"></a><span data-ttu-id="046c1-157">Вставка надстроек Office из каталога приложений</span><span class="sxs-lookup"><span data-stu-id="046c1-157">Insert Office Add-ins from the app catalog</span></span>

<span data-ttu-id="046c1-158">Для веб-приложений Office надстройки Office можно найти в каталоге приложений, выполнив следующие действия.</span><span class="sxs-lookup"><span data-stu-id="046c1-158">For online Office applications, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="046c1-159">Откройте веб-приложение Office (Excel, PowerPoint или Word).</span><span class="sxs-lookup"><span data-stu-id="046c1-159">Open the online Office application (Excel, PowerPoint, or Word).</span></span>
2. <span data-ttu-id="046c1-160">Создайте или откройте документ.</span><span class="sxs-lookup"><span data-stu-id="046c1-160">Create or open a document.</span></span>
3. <span data-ttu-id="046c1-161">Выберите **Вставка** > **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="046c1-161">Choose **Insert** > **Add-ins**.</span></span>
4. <span data-ttu-id="046c1-162">В диалоговом окне "Надстройки Office" выберите вкладку **МОЯ ОРГАНИЗАЦИЯ**. Отобразится список надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="046c1-162">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
5. <span data-ttu-id="046c1-163">Выберите надстройку Office и нажмите **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="046c1-163">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="046c1-164">Для классических приложений Office надстройки Office можно найти в каталоге приложений, выполнив следующие действия.</span><span class="sxs-lookup"><span data-stu-id="046c1-164">For Office applications on the desktop, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="046c1-165">Откройте классическое приложение Office (Excel, Word или PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="046c1-165">Open the desktop Office application (Excel, Word, or PowerPoint)</span></span>
2. <span data-ttu-id="046c1-166">Выберите **Файл** > **Параметры** > **Центр управления безопасностью** > **Параметры центра управления безопасностью** > **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="046c1-166">Choose **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
3. <span data-ttu-id="046c1-167">Введите URL-адрес каталога приложений SharePoint в поле **URL-адрес каталога** и нажмите кнопку **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="046c1-167">Enter the URL of the SharePoint app catalog in the **Catalog Url** box and choose **Add catalog**.</span></span>
    <span data-ttu-id="046c1-168">Используйте укороченный формат URL-адреса.</span><span class="sxs-lookup"><span data-stu-id="046c1-168">Use the shorter form of the URL.</span></span> <span data-ttu-id="046c1-169">Предположим, что URL-адрес каталога приложений SharePoint такой:</span><span class="sxs-lookup"><span data-stu-id="046c1-169">For example, if the URL of the Office Add-ins catalog is:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`
    
    <span data-ttu-id="046c1-170">Укажите только URL-адрес родительского семейства веб-сайтов:</span><span class="sxs-lookup"><span data-stu-id="046c1-170">Specify just the URL of the parent site collection:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
4. <span data-ttu-id="046c1-171">Закройте приложение Office и снова запустите его.</span><span class="sxs-lookup"><span data-stu-id="046c1-171">Close and reopen the Office application.</span></span> 
5. <span data-ttu-id="046c1-172">Выберите **Вставка** > **Получить надстройки**.</span><span class="sxs-lookup"><span data-stu-id="046c1-172">Choose **Insert** > **Get Add-ins**.</span></span>
4. <span data-ttu-id="046c1-173">В диалоговом окне "Надстройки Office" выберите вкладку **МОЯ ОРГАНИЗАЦИЯ**. Отобразится список надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="046c1-173">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
5. <span data-ttu-id="046c1-174">Выберите надстройку Office и нажмите **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="046c1-174">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="046c1-175">Кроме того, администратор может указать каталог приложений в SharePoint с помощью групповой политики.</span><span class="sxs-lookup"><span data-stu-id="046c1-175">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="046c1-176">Соответствующие параметры политики доступны в [файлах административных шаблонов (ADMX/ADML) для Office 365 профессиональный плюс, Office 2019 и Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) и находятся в папке **User Configuration\Policies\Administrative Templates\Microsoft Office 2016\Security Settings\Trust Center\Trusted Catalogs**.</span><span class="sxs-lookup"><span data-stu-id="046c1-176">The relevant policy settings are available in the [Administrative Template files (ADMX/ADML) for Office 365 ProPlus, Office 2019, and Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) and be found under **User Configuration\Policies\Administrative Templates\Microsoft Office 2016\Security Settings\Trust Center\Trusted Catalogs**.</span></span>
