---
title: Публикация надстроек области задач и контентных надстроек в каталоге приложений SharePoint
description: Чтобы предоставить доступ к надстройкам Office пользователям в организации, администраторы могут отправлять файлы манифестов надстроек Office в соответствующий каталог приложений.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: e9a600cd807379e9c55f2fc98bb4f2d71552058f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325307"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-app-catalog"></a><span data-ttu-id="bad2f-103">Публикация надстроек области задач и контентных надстроек в каталоге приложений SharePoint</span><span class="sxs-lookup"><span data-stu-id="bad2f-103">Publish task pane and content add-ins to a SharePoint app catalog</span></span>

<span data-ttu-id="bad2f-p101">Каталог приложений — это отдельное семейство веб-сайтов в веб-приложении SharePoint или клиенте SharePoint Online, в котором размещены библиотеки документов для надстроек Office и SharePoint. Администраторы могут отправлять в него файлы манифестов надстроек Office, чтобы пользователи в организации могли получить доступ к этим надстройкам. Когда администратор зарегистрирует каталог приложений как доверенный, пользователи смогут вставлять надстройки в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="bad2f-p101">An app catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the app catalog for their organization. When an administrator registers an app catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="bad2f-106">Каталоги приложений в SharePoint не поддерживают функции надстроек, реализованные в узле `VersionOverrides` [манифеста надстройки](../develop/add-in-manifests.md), такие как команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="bad2f-106">App catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="bad2f-107">Чтобы публиковать надстройки для облачной или гибридной среды, рекомендуем использовать [централизованное развертывание через Центр администрирования Office 365](../publish/centralized-deployment.md).</span><span class="sxs-lookup"><span data-stu-id="bad2f-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="bad2f-108">Каталоги приложений SharePoint не поддерживаются в Office для Mac.</span><span class="sxs-lookup"><span data-stu-id="bad2f-108">App catalogs on SharePoint are not supported in Office on Mac.</span></span> <span data-ttu-id="bad2f-109">Для развертывания надстроек Office на клиентах Mac необходимо отправить их в [AppSource](/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="bad2f-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="create-an-app-catalog"></a><span data-ttu-id="bad2f-110">Создание каталога приложений</span><span class="sxs-lookup"><span data-stu-id="bad2f-110">Create an app catalog</span></span>

<span data-ttu-id="bad2f-111">Выполните действия, описанные в одном из указанных ниже разделов, чтобы создать каталог приложений в локальном сервере SharePoint Server или Office 365.</span><span class="sxs-lookup"><span data-stu-id="bad2f-111">Complete the steps in one of the following sections to create an app catalog with on-premises SharePoint Server or on Office 365.</span></span>

### <a name="to-create-an-app-catalog-for-on-premises-sharepoint-server"></a><span data-ttu-id="bad2f-112">Создание каталога приложений в локальном сервере SharePoint Server</span><span class="sxs-lookup"><span data-stu-id="bad2f-112">To create an app catalog for on-premises SharePoint Server</span></span>

<span data-ttu-id="bad2f-113">Чтобы создать каталог приложений SharePoint, следуйте инструкциям в статье [Настройка сайта каталога приложений для веб-приложения](/sharepoint/administration/manage-the-app-catalog).</span><span class="sxs-lookup"><span data-stu-id="bad2f-113">To create the SharePoint app catalog, follow the instructions at [Configure the App Catalog site for a web application](/sharepoint/administration/manage-the-app-catalog).</span></span>

<span data-ttu-id="bad2f-114">После создания каталога приложений выполните инструкции [по публикации надстройки Office](#publish-an-office-add-in).</span><span class="sxs-lookup"><span data-stu-id="bad2f-114">Once you have created the app catalog follow the steps to [publish an Office Add-in](#publish-an-office-add-in).</span></span>

### <a name="to-create-an-app-catalog-on-office-365"></a><span data-ttu-id="bad2f-115">Создание каталога приложений в Office 365</span><span class="sxs-lookup"><span data-stu-id="bad2f-115">To create an app catalog on Office 365</span></span>

1. <span data-ttu-id="bad2f-116">Перейдите в Центр администрирования Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="bad2f-116">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="bad2f-117">Сведения о том, как найти Центр администрирования, см. в статье [Сведения о Центре администрирования Microsoft 365](/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="bad2f-117">For information on how to find the admin center, see [About the Microsoft 365 admin center](/office365/admin/admin-overview/about-the-admin-center).</span></span>

2. <span data-ttu-id="bad2f-118">На странице Центра администрирования Microsoft 365 разверните список **центров администрирования** и выберите пункт **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-118">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="bad2f-119">Для создания каталога нужно использовать классический Центр администрирования SharePoint.</span><span class="sxs-lookup"><span data-stu-id="bad2f-119">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="bad2f-120">Если вы находитесь в новом Центре администрирования SharePoint, выберите пункт **Классический Центр администрирования SharePoint** в области слева.</span><span class="sxs-lookup"><span data-stu-id="bad2f-120">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>

3. <span data-ttu-id="bad2f-121">В области задач слева выберите пункт **приложения**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-121">In the left task pane, choose **apps**.</span></span>

4. <span data-ttu-id="bad2f-122">На странице **приложения** выберите пункт **Каталог приложений**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-122">On the **apps** page, choose **App Catalog**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="bad2f-123">Если каталог приложений уже создан и отображается на этой странице, вы можете пропустить остальные этапы и перейти к следующему разделу этой статьи, чтобы опубликовать надстройку в каталоге.</span><span class="sxs-lookup"><span data-stu-id="bad2f-123">If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.</span></span>

5. <span data-ttu-id="bad2f-124">На странице **Сайт каталога приложений** нажмите кнопку **ОК**, чтобы принять параметр по умолчанию и создать сайт каталога приложений.</span><span class="sxs-lookup"><span data-stu-id="bad2f-124">On the **App Catalog Site** page, choose **OK** to accept the default option and create a new app catalog site.</span></span>

6. <span data-ttu-id="bad2f-125">На странице **Создание семейства веб-сайтов каталога приложений** укажите название сайта каталога приложений.</span><span class="sxs-lookup"><span data-stu-id="bad2f-125">On the **Create App Catalog Site Collection** page, specify the title of your App Catalog site.</span></span>

7. <span data-ttu-id="bad2f-126">Укажите **адрес веб-сайта**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-126">Specify the **Web Site Address**.</span></span>

8. <span data-ttu-id="bad2f-127">Укажите **администратора**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-127">Specify an **Administrator**.</span></span>

9. <span data-ttu-id="bad2f-128">Для параметра **Квота ресурсов сервера** установите значение 0 (ноль).</span><span class="sxs-lookup"><span data-stu-id="bad2f-128">Set the **Server Resource Quota** to 0 (zero).</span></span> <span data-ttu-id="bad2f-129">(Квота ресурсов сервера связана с регулированием изолированных решений с низкой производительностью, но вы не будете устанавливать изолированные решения на сайте каталога приложений.)</span><span class="sxs-lookup"><span data-stu-id="bad2f-129">(The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your app catalog site.)</span></span>

10. <span data-ttu-id="bad2f-130">Нажмите кнопку **OK**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-130">Choose **OK**.</span></span>

## <a name="publish-an-office-add-in"></a><span data-ttu-id="bad2f-131">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="bad2f-131">Publish an Office Add-in</span></span>

<span data-ttu-id="bad2f-132">Выполните действия, описанные в одном из указанных ниже разделов, чтобы опубликовать надстройку Office в каталоге приложений в Office 365 или локальном сервере SharePoint Server.</span><span class="sxs-lookup"><span data-stu-id="bad2f-132">Complete the steps in one of the following sections to publish an Office Add-in to an app catalog on Office 365 or on-premises SharePoint Server.</span></span>

### <a name="to-publish-an-office-add-in-to-a-sharepoint-app-catalog-on-office-365"></a><span data-ttu-id="bad2f-133">Публикация надстройки Office в каталоге приложений SharePoint в Office 365</span><span class="sxs-lookup"><span data-stu-id="bad2f-133">To publish an Office add-in to a SharePoint app catalog on Office 365</span></span>

1. <span data-ttu-id="bad2f-134">Перейдите в Центр администрирования Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="bad2f-134">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="bad2f-135">Сведения о том, как найти Центр администрирования, см. в статье [Сведения о Центре администрирования Microsoft 365](/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="bad2f-135">For information on how to find the admin center, see [About the Microsoft 365 admin center](/office365/admin/admin-overview/about-the-admin-center).</span></span>
2. <span data-ttu-id="bad2f-136">На странице Центра администрирования Microsoft 365 разверните список **центров администрирования** и выберите пункт **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-136">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="bad2f-137">Для создания каталога нужно использовать классический Центр администрирования SharePoint.</span><span class="sxs-lookup"><span data-stu-id="bad2f-137">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="bad2f-138">Если вы находитесь в новом Центре администрирования SharePoint, выберите пункт **Классический Центр администрирования SharePoint** в области слева.</span><span class="sxs-lookup"><span data-stu-id="bad2f-138">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>
3. <span data-ttu-id="bad2f-139">В области задач слева выберите пункт **приложения**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-139">In the left task pane, choose **apps**.</span></span>
4. <span data-ttu-id="bad2f-140">На странице **приложения** выберите пункт **Каталог приложений**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-140">On the **apps** page, choose **App Catalog**.</span></span>
5. <span data-ttu-id="bad2f-141">Выберите элемент **Распределить приложения для Office**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-141">Choose **Distribute apps for Office**.</span></span>
6. <span data-ttu-id="bad2f-142">На странице **Приложения для Office** выберите команду **Создать**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-142">In the **Apps for Office** page, choose **New**.</span></span>
7. <span data-ttu-id="bad2f-143">В диалоговом окне **Добавление документа** нажмите кнопку **Выбрать файлы**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-143">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
8. <span data-ttu-id="bad2f-144">Найдите и укажите файл [манифеста](../develop/add-in-manifests.md) для добавления и нажмите кнопку **Открыть**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-144">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
9. <span data-ttu-id="bad2f-145">В диалоговом окне **Добавление документа** нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-145">In the **Add a document** dialog, choose **OK**.</span></span>

### <a name="to-publish-an-add-in-to-an-app-catalog-with-on-premises-sharepoint-server"></a><span data-ttu-id="bad2f-146">Публикация надстройки в каталоге приложений с помощью локального сервера SharePoint Server</span><span class="sxs-lookup"><span data-stu-id="bad2f-146">To publish an add-in to an app catalog with on-premises SharePoint Server</span></span>

1. <span data-ttu-id="bad2f-147">Откройте страницу **Центр администрирования**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-147">Open the **Central Administration** page.</span></span>
2. <span data-ttu-id="bad2f-148">В области задач слева выберите пункт **Приложения**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-148">In the left task pane, choose **Apps**.</span></span>
3. <span data-ttu-id="bad2f-149">На странице **Приложения** в разделе **Управление приложениями** выберите пункт **Управление каталогом приложений**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-149">On the **Apps** page, under **App Management**, choose **Manage App Catalog**.</span></span>
4. <span data-ttu-id="bad2f-150">На странице **Управление каталогом приложений** убедитесь, что в поле выбора **Веб-приложение** выбрано правильное веб-приложение.</span><span class="sxs-lookup"><span data-stu-id="bad2f-150">On the **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application** Selector.</span></span>
5. <span data-ttu-id="bad2f-151">Выберите URL-адрес в разделе **URL-адрес сайта**, чтобы открыть сайт каталога приложений.</span><span class="sxs-lookup"><span data-stu-id="bad2f-151">Choose the URL under the **Site URL** to open the app catalog site.</span></span>
6. <span data-ttu-id="bad2f-152">Выберите элемент **Распределить приложения для Office**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-152">Choose **Distribute apps for Office**.</span></span>
7. <span data-ttu-id="bad2f-153">На странице **Приложения для Office** выберите команду **Создать**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-153">In the **Apps for Office** page, choose **New**.</span></span>
8. <span data-ttu-id="bad2f-154">В диалоговом окне **Добавление документа** нажмите кнопку **Выбрать файлы**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-154">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
9. <span data-ttu-id="bad2f-155">Найдите и укажите файл [манифеста](../develop/add-in-manifests.md) для добавления и нажмите кнопку **Открыть**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-155">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
10. <span data-ttu-id="bad2f-156">В диалоговом окне **Добавление документа** нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-156">In the **Add a document** dialog, choose **OK**.</span></span>

## <a name="insert-office-add-ins-from-the-app-catalog"></a><span data-ttu-id="bad2f-157">Вставка надстроек Office из каталога приложений</span><span class="sxs-lookup"><span data-stu-id="bad2f-157">Insert Office Add-ins from the app catalog</span></span>

<span data-ttu-id="bad2f-158">Для веб-приложений Office надстройки Office можно найти в каталоге приложений, выполнив следующие действия.</span><span class="sxs-lookup"><span data-stu-id="bad2f-158">For online Office applications, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="bad2f-159">Откройте веб-приложение Office (Excel, PowerPoint или Word).</span><span class="sxs-lookup"><span data-stu-id="bad2f-159">Open the online Office application (Excel, PowerPoint, or Word).</span></span>
2. <span data-ttu-id="bad2f-160">Создайте или откройте документ.</span><span class="sxs-lookup"><span data-stu-id="bad2f-160">Create or open a document.</span></span>
3. <span data-ttu-id="bad2f-161">Выберите **Вставка** > **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-161">Choose **Insert** > **Add-ins**.</span></span>
4. <span data-ttu-id="bad2f-162">В диалоговом окне "Надстройки Office" выберите вкладку **МОЯ ОРГАНИЗАЦИЯ**. Отобразится список надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="bad2f-162">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
5. <span data-ttu-id="bad2f-163">Выберите надстройку Office и нажмите **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-163">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="bad2f-164">Для классических приложений Office надстройки Office можно найти в каталоге приложений, выполнив следующие действия.</span><span class="sxs-lookup"><span data-stu-id="bad2f-164">For Office applications on the desktop, you can find Office Add-ins from the app catalog by completing the following steps.</span></span>

1. <span data-ttu-id="bad2f-165">Откройте классическое приложение Office (Excel, Word или PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="bad2f-165">Open the desktop Office application (Excel, Word, or PowerPoint)</span></span>
2. <span data-ttu-id="bad2f-166">Выберите **Файл** > **Параметры** > **Центр управления безопасностью** > **Параметры центра управления безопасностью** > **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-166">Choose **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
3. <span data-ttu-id="bad2f-167">Введите URL-адрес каталога приложений SharePoint в поле **URL-адрес каталога** и нажмите кнопку **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-167">Enter the URL of the SharePoint app catalog in the **Catalog Url** box and choose **Add catalog**.</span></span>
    <span data-ttu-id="bad2f-168">Используйте укороченный формат URL-адреса.</span><span class="sxs-lookup"><span data-stu-id="bad2f-168">Use the shorter form of the URL.</span></span> <span data-ttu-id="bad2f-169">Предположим, что URL-адрес каталога приложений SharePoint такой:</span><span class="sxs-lookup"><span data-stu-id="bad2f-169">For example, if the URL of the SharePoint app catalog is:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>/AgaveCatalog`
    
    <span data-ttu-id="bad2f-170">Укажите только URL-адрес родительского семейства веб-сайтов:</span><span class="sxs-lookup"><span data-stu-id="bad2f-170">Specify just the URL of the parent site collection:</span></span>
    - `https://<domain>/sites/<AddinCatalogSiteCollection>`
4. <span data-ttu-id="bad2f-171">Закройте приложение Office и снова запустите его.</span><span class="sxs-lookup"><span data-stu-id="bad2f-171">Close and reopen the Office application.</span></span> 
5. <span data-ttu-id="bad2f-172">Выберите **Вставка** > **Получить надстройки**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-172">Choose **Insert** > **Get Add-ins**.</span></span>
4. <span data-ttu-id="bad2f-173">В диалоговом окне "Надстройки Office" выберите вкладку **МОЯ ОРГАНИЗАЦИЯ**. Отобразится список надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="bad2f-173">In the Office Add-ins dialog, choose the **MY ORGANIZATION** tab.  The Office Add-ins are listed.</span></span>
5. <span data-ttu-id="bad2f-174">Выберите надстройку Office и нажмите **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-174">Choose an Office Add-in and then choose **Add**.</span></span>

<span data-ttu-id="bad2f-175">Кроме того, администратор может указать каталог приложений в SharePoint с помощью групповой политики.</span><span class="sxs-lookup"><span data-stu-id="bad2f-175">Alternatively, an administrator can specify an app catalog on SharePoint by using Group Policy.</span></span> <span data-ttu-id="bad2f-176">Соответствующие параметры политики доступны в [файлах административных шаблонов (ADMX/ADML) для Office 365 профессиональный плюс, Office 2019 и Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) и находятся в папке **User Configuration\Policies\Administrative Templates\Microsoft Office 2016\Security Settings\Trust Center\Trusted Catalogs**.</span><span class="sxs-lookup"><span data-stu-id="bad2f-176">The relevant policy settings are available in the [Administrative Template files (ADMX/ADML) for Office 365 ProPlus, Office 2019, and Office 2016](https://www.microsoft.com/download/details.aspx?id=49030) and be found under **User Configuration\Policies\Administrative Templates\Microsoft Office 2016\Security Settings\Trust Center\Trusted Catalogs**.</span></span>
