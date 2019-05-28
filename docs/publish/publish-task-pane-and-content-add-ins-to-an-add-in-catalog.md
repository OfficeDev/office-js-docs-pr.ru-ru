---
title: Публикация надстроек области задач и контентных надстроек в каталоге SharePoint
description: Чтобы делать надстройки Office доступными пользователям в организации, администраторы могут отправлять файлы манифестов надстроек Office в соответствующий каталог надстроек.
ms.date: 05/22/2019
localization_priority: Priority
ms.openlocfilehash: bffbf3e83a2e6d8d0c63252c27ba54826611f78b
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432245"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="6e25b-103">Публикация надстроек области задач и контентных надстроек в каталоге SharePoint</span><span class="sxs-lookup"><span data-stu-id="6e25b-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="6e25b-p101">Каталог надстроек — это отдельное семейство веб-сайтов в веб-приложении SharePoint или клиенте SharePoint Online, в котором размещены библиотеки документов для надстроек Office и SharePoint. Администраторы могут отправлять в него файлы манифестов надстроек Office, чтобы пользователи в организации могли получить доступ к этим надстройкам. Когда администратор зарегистрирует каталог надстроек как доверенный, пользователи смогут вставлять надстройки в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="6e25b-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="6e25b-106">Каталоги надстроек в SharePoint не поддерживают функции надстроек, реализованные в узле `VersionOverrides` [манифеста надстройки](../develop/add-in-manifests.md), такие как команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="6e25b-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="6e25b-107">Чтобы публиковать надстройки для облачной или гибридной среды, рекомендуем использовать [централизованное развертывание через Центр администрирования Office 365](../publish/centralized-deployment.md).</span><span class="sxs-lookup"><span data-stu-id="6e25b-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="6e25b-108">Каталоги SharePoint не поддерживаются в Office для Mac.</span><span class="sxs-lookup"><span data-stu-id="6e25b-108">SharePoint catalogs are not supported for Office for Mac.</span></span> <span data-ttu-id="6e25b-109">Для развертывания надстроек Office на клиентах Mac необходимо отправить их в [AppSource](/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="6e25b-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="create-an-add-in-catalog"></a><span data-ttu-id="6e25b-110">Создание каталога надстроек</span><span class="sxs-lookup"><span data-stu-id="6e25b-110">Create an add-in catalog</span></span>

<span data-ttu-id="6e25b-111">Выполните действия, описанные в одном из указанных ниже разделов, чтобы создать каталог надстроек в SharePoint или Office 365.</span><span class="sxs-lookup"><span data-stu-id="6e25b-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-create-an-add-in-catalog-for-on-premises-sharepoint"></a><span data-ttu-id="6e25b-112">Создание каталога надстроек в локальном SharePoint</span><span class="sxs-lookup"><span data-stu-id="6e25b-112">To set up an add-in catalog for on-premises SharePoint</span></span>

> [!NOTE]
> <span data-ttu-id="6e25b-113">Надстройки в пользовательском интерфейсе локального SharePoint по-прежнему называются **приложениями**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-113">The UI in on-premises SharePoint still refers to add-ins as **apps**.</span></span>

1. <span data-ttu-id="6e25b-114">Перейдите на **сайт центра администрирования**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-114">Browse to the  **Central Administration Site**.</span></span>

2. <span data-ttu-id="6e25b-115">В области задач слева выберите пункт **Приложения**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-115">In the left task pane, choose  **Apps**.</span></span>

3. <span data-ttu-id="6e25b-116">На странице **Приложения** в разделе **Управление приложениями** выберите пункт **Управление каталогом приложений**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-116">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>

4. <span data-ttu-id="6e25b-117">На странице  **Управление каталогом приложений** убедитесь, что в пункте **Селектор веб-приложения** выбрано правильное веб-приложение.</span><span class="sxs-lookup"><span data-stu-id="6e25b-117">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>

5. <span data-ttu-id="6e25b-118">Выберите элемент **Просмотреть параметры сайта**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-118">Choose  **View site settings**.</span></span>

6. <span data-ttu-id="6e25b-119">На странице  **Параметры сайта** выберите пункт **Администраторы семейства веб-сайтов**, чтобы указать администраторов семейства веб-сайтов, а затем нажмите кнопку  **ОК**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-119">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>

7. <span data-ttu-id="6e25b-120">Чтобы предоставить пользователям разрешения для сайтов, последовательно выберите элементы  **Разрешения для сайта** и **Предоставить разрешения**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-120">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>

8. <span data-ttu-id="6e25b-121">В диалоговом окне  **Общий доступ к сайту каталога приложений** укажите одного или нескольких пользователей сайта, задайте соответствующие разрешения для них, при необходимости укажите другие параметры, а затем выберите элемент **Общий доступ**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-121">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>

9. <span data-ttu-id="6e25b-122">Чтобы добавить надстройку в каталог надстроек Office, выберите **Приложения для Office**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-122">To add an add-in to the Office Add-ins add-in catalog, choose **Apps for Office**.</span></span>

### <a name="to-create-an-app-catalog-on-office-365"></a><span data-ttu-id="6e25b-123">Создание каталога приложений в Office 365</span><span class="sxs-lookup"><span data-stu-id="6e25b-123">To create an app catalog on Office 365</span></span>

<span data-ttu-id="6e25b-124">Хотя SharePoint называет его каталогом "приложений", вы можете регистрировать надстройки Office в каталоге приложений.</span><span class="sxs-lookup"><span data-stu-id="6e25b-124">Even though SharePoint names the catalog an "app" catalog, you can register Office Add-ins in the app catalog.</span></span>

1. <span data-ttu-id="6e25b-125">Перейдите в Центр администрирования Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="6e25b-125">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="6e25b-126">Сведения о том, как найти Центр администрирования, см. в статье [Сведения о Центре администрирования Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="6e25b-126">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>

2. <span data-ttu-id="6e25b-127">На странице Центра администрирования Microsoft 365 разверните список **центров администрирования** и выберите пункт **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-127">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6e25b-128">Для создания каталога нужно использовать классический Центр администрирования SharePoint.</span><span class="sxs-lookup"><span data-stu-id="6e25b-128">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="6e25b-129">Если вы находитесь в новом Центре администрирования SharePoint, выберите пункт **Классический Центр администрирования SharePoint** в области слева.</span><span class="sxs-lookup"><span data-stu-id="6e25b-129">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>

3. <span data-ttu-id="6e25b-130">В области задач слева выберите пункт **приложения**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-130">In the left task pane, choose  **Apps**.</span></span>

4. <span data-ttu-id="6e25b-131">На странице **приложения** выберите пункт **Каталог приложений**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-131">On the **apps** page, select **App Catalog**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="6e25b-132">Если каталог приложений уже создан и отображается на этой странице, вы можете пропустить остальные этапы и перейти к следующему разделу этой статьи, чтобы опубликовать надстройку в каталоге.</span><span class="sxs-lookup"><span data-stu-id="6e25b-132">If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.</span></span>

5. <span data-ttu-id="6e25b-133">На странице **Сайт каталога приложений** нажмите кнопку **ОК**, чтобы принять параметр по умолчанию и создать сайт каталога надстроек.</span><span class="sxs-lookup"><span data-stu-id="6e25b-133">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>

6. <span data-ttu-id="6e25b-134">На странице **Создание семейства веб-сайтов каталога приложений** укажите название сайта каталога приложений.</span><span class="sxs-lookup"><span data-stu-id="6e25b-134">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

7. <span data-ttu-id="6e25b-135">Укажите **адрес веб-сайта**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-135">Specify the web site address.</span></span>

8. <span data-ttu-id="6e25b-136">Укажите **администратора**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-136">Specify an **Administrator**.</span></span>

9. <span data-ttu-id="6e25b-137">Для параметра **Квота ресурсов сервера** установите значение 0 (ноль).</span><span class="sxs-lookup"><span data-stu-id="6e25b-137">Set the **Server Resource Quota** to 0 (zero).</span></span> <span data-ttu-id="6e25b-138">(Квота ресурсов сервера связана с регулированием изолированных решений с низкой производительностью, но вы не будете устанавливать изолированные решения на сайте каталога приложений.)</span><span class="sxs-lookup"><span data-stu-id="6e25b-138">(The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

10. <span data-ttu-id="6e25b-139">Нажмите кнопку **OK**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-139">Choose **OK**.</span></span>

<span data-ttu-id="6e25b-140">Каталог приложений создан.</span><span class="sxs-lookup"><span data-stu-id="6e25b-140">The app catalog is now created.</span></span>

## <a name="publish-an-add-in-to-an-app-catalog"></a><span data-ttu-id="6e25b-141">Публикация надстройки в каталоге приложений</span><span class="sxs-lookup"><span data-stu-id="6e25b-141">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="6e25b-142">Чтобы опубликовать надстройку в существующем каталоге приложений, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="6e25b-142">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="6e25b-143">Перейдите в Центр администрирования Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="6e25b-143">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="6e25b-144">Сведения о том, как найти Центр администрирования, см. в статье [Сведения о Центре администрирования Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="6e25b-144">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>
2. <span data-ttu-id="6e25b-145">На странице Центра администрирования Microsoft 365 разверните список **центров администрирования** и выберите пункт **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-145">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="6e25b-146">Для создания каталога нужно использовать классический Центр администрирования SharePoint.</span><span class="sxs-lookup"><span data-stu-id="6e25b-146">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="6e25b-147">Если вы находитесь в новом Центре администрирования SharePoint, выберите пункт **Классический Центр администрирования SharePoint** в области слева.</span><span class="sxs-lookup"><span data-stu-id="6e25b-147">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>
3. <span data-ttu-id="6e25b-148">В области задач слева выберите пункт **приложения**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-148">In the left task pane, choose  **Apps**.</span></span>
4. <span data-ttu-id="6e25b-149">На странице **приложения** выберите пункт **Каталог приложений**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-149">On the **apps** page, select **App Catalog**.</span></span>
5. <span data-ttu-id="6e25b-150">Выберите элемент **Распределить приложения для Office**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-150">Choose **Distribute apps for Office**.</span></span>
6. <span data-ttu-id="6e25b-151">На странице **Приложения для Office** выберите команду **Создать**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-151">In the **Apps for Office** page, choose **New**.</span></span>
7. <span data-ttu-id="6e25b-152">В диалоговом окне **Добавление документа** нажмите кнопку **Выбрать файлы**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-152">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
8. <span data-ttu-id="6e25b-153">Найдите и укажите файл [манифеста](../develop/add-in-manifests.md) для добавления и нажмите кнопку **Открыть**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-153">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
9. <span data-ttu-id="6e25b-154">В диалоговом окне **Добавление документа** нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-154">In the **Add a document** dialog box, choose **OK**.</span></span>

    <span data-ttu-id="6e25b-p108">Теперь надстройки области задач и контентные надстройки из этого каталога доступны в диалоговом окне **Надстройки Office**. Для доступа к ним выберите**Мои надстройки** на вкладке **Вставка**, а затем нажмите **Моя организация**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-p108">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="6e25b-157">Работа пользователей с каталогом надстроек</span><span class="sxs-lookup"><span data-stu-id="6e25b-157">End user experience with the add-in catalog</span></span>

<span data-ttu-id="6e25b-158">Пользователь может получить доступ к каталогу надстроек в приложении Office, выполнив указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="6e25b-158">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="6e25b-159">В приложении Office выберите **Файл** > **Параметры** > **Центр управления безопасностью** > **Параметры центра управления безопасностью** > **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-159">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>

2. <span data-ttu-id="6e25b-160">Укажите URL-адрес _родительского семейства веб-сайтов SharePoint_ для каталога надстроек.</span><span class="sxs-lookup"><span data-stu-id="6e25b-160">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 

    <span data-ttu-id="6e25b-161">Предположим, что URL-адрес каталога надстроек Office такой:</span><span class="sxs-lookup"><span data-stu-id="6e25b-161">For example, if the URL of the Office Add-ins catalog is:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`

    <span data-ttu-id="6e25b-162">Укажите только URL-адрес родительского семейства веб-сайтов:</span><span class="sxs-lookup"><span data-stu-id="6e25b-162">Specify just the URL of the parent site collection:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`

3. <span data-ttu-id="6e25b-p109">Закройте приложение Office и снова запустите его. Каталог надстроек будет доступен в диалоговом окне **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="6e25b-p109">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="6e25b-165">Кроме того, администратор может указать каталог надстроек Office в SharePoint с помощью групповой политики.</span><span class="sxs-lookup"><span data-stu-id="6e25b-165">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="6e25b-166">Дополнительные сведения см. в разделе [Использование групповой политики для управления возможностью установки и использования пользователями приложений для Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span><span class="sxs-lookup"><span data-stu-id="6e25b-166">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span></span>
