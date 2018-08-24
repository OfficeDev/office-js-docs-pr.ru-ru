---
title: Публикация надстроек области задач и контентных надстроек в каталоге SharePoint
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 6bf63c36d952b901faaa16b0d93748023ac0fef9
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925299"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="b8642-102">Публикация надстроек области задач и контентных надстроек в каталоге SharePoint</span><span class="sxs-lookup"><span data-stu-id="b8642-102">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="b8642-p101">Каталог надстроек — это отдельное семейство веб-сайтов в веб-приложении SharePoint или клиенте SharePoint Online, в котором размещены библиотеки документов для надстроек Office и SharePoint. Администраторы могут отправлять в него файлы манифестов надстроек Office, чтобы пользователи в организации могли получить доступ к этим надстройкам. Когда администратор зарегистрирует каталог надстроек как доверенный, пользователи смогут вставлять надстройки в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="b8642-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="b8642-105">Каталоги надстроек в SharePoint не поддерживают функции надстроек, реализованные в узле `VersionOverrides` [манифеста надстройки](../develop/add-in-manifests.md), такие как команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="b8642-105">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="b8642-106">Чтобы публиковать надстройки для облачной или гибридной среды, рекомендуем использовать [централизованное развертывание через Центр администрирования Office 365](../publish/centralized-deployment.md).</span><span class="sxs-lookup"><span data-stu-id="b8642-106">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="b8642-p102">Каталоги SharePoint не поддерживаются в Office 2016 для Mac. Для развертывания надстроек Office на клиентах Mac нужно отправить их в [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="b8642-p102">SharePoint catalogs are not supported for Office 2016 for Mac. To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="set-up-an-add-in-catalog"></a><span data-ttu-id="b8642-109">Настройка каталога надстроек</span><span class="sxs-lookup"><span data-stu-id="b8642-109">Set up an add-in catalog</span></span>

<span data-ttu-id="b8642-110">Выполните действия, описанные в одном из указанных ниже разделов, чтобы настроить каталог надстроек в SharePoint или Office 365.</span><span class="sxs-lookup"><span data-stu-id="b8642-110">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-set-up-an-add-in-catalog-on-sharepoint"></a><span data-ttu-id="b8642-111">Настройка каталога надстроек в SharePoint</span><span class="sxs-lookup"><span data-stu-id="b8642-111">To set up an add-in catalog on SharePoint</span></span>

1. <span data-ttu-id="b8642-112">Откройте **сайт центра администрирования** (**Пуск** > **Все программы** > **Продукты Microsoft SharePoint 2013** > **Центр администрирования SharePoint 2013**).</span><span class="sxs-lookup"><span data-stu-id="b8642-112">Browse to the  **Central Administration Site** ( **Start** > **All Programs** > **Microsoft SharePoint 2013 Products** > **SharePoint 2013 Central Administration**).</span></span>
    
2. <span data-ttu-id="b8642-113">В области задач слева выберите пункт  **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="b8642-113">In the left task pane, choose  **Add-ins**.</span></span>
    
3. <span data-ttu-id="b8642-114">На странице  **Надстройки**, в разделе  **Управление надстройками**, выберите пункт  **Управление каталогом надстроек**.</span><span class="sxs-lookup"><span data-stu-id="b8642-114">On the  **Add-ins** page, under **Add-in Management**, choose  **Manage Add-in Catalog**.</span></span>
    
4. <span data-ttu-id="b8642-115">На странице  **Управление каталогом надстроек** убедитесь, что в пункте **Селектор веб-приложения** выбрано правильное веб-приложение.</span><span class="sxs-lookup"><span data-stu-id="b8642-115">On the  **Manage Add-in Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>
    
5. <span data-ttu-id="b8642-116">Выберите элемент  **Просмотреть параметры сайта**.</span><span class="sxs-lookup"><span data-stu-id="b8642-116">Choose  **View site settings**.</span></span>
    
6. <span data-ttu-id="b8642-117">На странице  **Параметры сайта** выберите пункт **Администраторы семейства веб-сайтов**, чтобы указать администраторов семейства веб-сайтов, а затем нажмите кнопку  **ОК**.</span><span class="sxs-lookup"><span data-stu-id="b8642-117">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>
    
7. <span data-ttu-id="b8642-118">Чтобы предоставить пользователям разрешения для сайтов, последовательно выберите элементы  **Разрешения для сайта** и **Предоставить разрешения**.</span><span class="sxs-lookup"><span data-stu-id="b8642-118">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>
    
8. <span data-ttu-id="b8642-119">В диалоговом окне  **Общий доступ к сайту каталога приложений** укажите одного или нескольких пользователей сайта, задайте соответствующие разрешения для них, при необходимости укажите другие параметры, а затем выберите элемент **Общий доступ**.</span><span class="sxs-lookup"><span data-stu-id="b8642-119">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>
    
9. <span data-ttu-id="b8642-120">Чтобы добавить надстройку в каталог надстроек Office, выберите **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="b8642-120">To add an add-in to the Office Add-ins add-in catalog, choose **Office Add-ins**.</span></span>

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a><span data-ttu-id="b8642-121">Настройка каталога надстроек в Office 365</span><span class="sxs-lookup"><span data-stu-id="b8642-121">To set up an add-in catalog on Office 365</span></span>

1. <span data-ttu-id="b8642-122">На странице Центра администрирования Office 365 выберите элемент **Администратор**, а затем **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="b8642-122">On the Office 365 admin center page, choose  **Admin**, and then choose  **SharePoint**.</span></span>
    
2. <span data-ttu-id="b8642-123">В области задач слева выберите пункт  **надстройки**.</span><span class="sxs-lookup"><span data-stu-id="b8642-123">In the left task pane, choose  **add-ins**.</span></span>
    
3. <span data-ttu-id="b8642-124">На странице  **надстройки** выберите пункт **Каталог надстроек**.</span><span class="sxs-lookup"><span data-stu-id="b8642-124">On the  **add-ins** page, choose **Add-in Catalog**.</span></span>
    
4. <span data-ttu-id="b8642-125">На странице  **Сайт каталога надстроек** нажмите кнопку **ОК**, чтобы принять параметр по умолчанию и создать сайт каталога надстроек.</span><span class="sxs-lookup"><span data-stu-id="b8642-125">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>
    
5. <span data-ttu-id="b8642-126">На странице  **Создание семейства веб-сайтов каталога надстроек** укажите название сайта каталога надстроек.</span><span class="sxs-lookup"><span data-stu-id="b8642-126">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>
    
6. <span data-ttu-id="b8642-127">Укажите адрес веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="b8642-127">Specify the web site address.</span></span>
    
7. <span data-ttu-id="b8642-p103">Минимальное допустимое значение (в данный момент оно составляет 110) указано в параметре  **Дисковая квота**. В этом семействе веб-сайтов будут устанавливаться только пакеты надстройка, которые имеют небольшой размер.</span><span class="sxs-lookup"><span data-stu-id="b8642-p103">Set the  **Storage Quota** to the lowest possible value (currently 110). You will only be installing add-in packages on this site collection and they are very small.</span></span>
    
8. <span data-ttu-id="b8642-p104">Задайте для параметра  **Квота ресурсов сервера** значение 0 (ноль). (Квота ресурсов сервера связана с регулированием изолированных решений с низкой производительностью, но на сайте каталога надстроек не будут устанавливаться изолированные решения.)</span><span class="sxs-lookup"><span data-stu-id="b8642-p104">Set the  **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>
    
9. <span data-ttu-id="b8642-132">Нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="b8642-132">Choose  **OK**.</span></span>
    
10. <span data-ttu-id="b8642-p105">Чтобы добавить надстройку на сайт каталога надстроек, перейдите на только что созданный сайт. В области навигации слева выберите пункт **Надстройки для Office**, а затем выберите команду **новая надстройка**, чтобы отправить надстройку для файла манифеста Office.</span><span class="sxs-lookup"><span data-stu-id="b8642-p105">To add an add-in to the Add-in Catalog Site, browse to the site you have just created. In the left navigation pane, choose  **Office Add-ins**, and then, to upload an Office Add-in manifest file, choose  **new add-in**.</span></span>

## <a name="publish-an-add-in-to-an-add-in-catalog"></a><span data-ttu-id="b8642-135">Публикация надстройки в каталоге надстроек</span><span class="sxs-lookup"><span data-stu-id="b8642-135">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="b8642-136">Чтобы опубликовать надстройку в каталоге надстроек, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="b8642-136">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="b8642-137">Перейдите в каталог надстроек.</span><span class="sxs-lookup"><span data-stu-id="b8642-137">Browse to the add-in catalog:</span></span>

    - <span data-ttu-id="b8642-138">Откройте главную страницу центра администрирования SharePoint.</span><span class="sxs-lookup"><span data-stu-id="b8642-138">Open the SharePoint Central Administration main page.</span></span>
    
    - <span data-ttu-id="b8642-139">Выберите **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="b8642-139">Select  **Add-ins**.</span></span>
    
    - <span data-ttu-id="b8642-140">Выберите **Управление каталогом надстроек**.</span><span class="sxs-lookup"><span data-stu-id="b8642-140">Select  **Manage Add-in Catalog**.</span></span>
    
    - <span data-ttu-id="b8642-141">Выберите указанную ссылку, а затем нажмите **Надстройки Office** на левой панели навигации.</span><span class="sxs-lookup"><span data-stu-id="b8642-141">Choose the link provided, and then choose  **Office Add-ins** on the left navigation bar.</span></span>
    
2. <span data-ttu-id="b8642-142">Выберите ссылку **Щелкните для добавления нового элемента**.</span><span class="sxs-lookup"><span data-stu-id="b8642-142">Choose the  **Click to add new item** link.</span></span>
    
3. <span data-ttu-id="b8642-143">Нажмите кнопку **Обзор**, а затем укажите [манифест](../develop/add-in-manifests.md) для отправки.</span><span class="sxs-lookup"><span data-stu-id="b8642-143">Choose  **Browse**, and then specify the [manifest](../develop/add-in-manifests.md) to upload.</span></span>
    
    <span data-ttu-id="b8642-p106">Теперь надстройки области задач и контентные надстройки из этого каталога доступны в диалоговом окне **Надстройки Office**. Для доступа к ним выберите**Мои надстройки** на вкладке **Вставка**, а затем нажмите **Моя организация**.</span><span class="sxs-lookup"><span data-stu-id="b8642-p106">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="b8642-146">Работа пользователей с каталогом надстроек</span><span class="sxs-lookup"><span data-stu-id="b8642-146">End user experience with the add-in catalog</span></span>

<span data-ttu-id="b8642-147">Пользователь может получить доступ к каталогу надстроек в приложении Office, выполнив указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="b8642-147">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="b8642-148">В приложении Office выберите **Файл** > **Параметры** > **Центр управления безопасностью** > **Параметры центра управления безопасностью** > **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="b8642-148">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
    
2. <span data-ttu-id="b8642-149">Укажите URL-адрес _родительского семейства веб-сайтов SharePoint_ для каталога надстроек.</span><span class="sxs-lookup"><span data-stu-id="b8642-149">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 
    
    <span data-ttu-id="b8642-150">Предположим, что URL-адрес каталога надстроек Office такой:</span><span class="sxs-lookup"><span data-stu-id="b8642-150">For example, if the URL of the Office Add-ins catalog is:</span></span>
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    <span data-ttu-id="b8642-151">Укажите только URL-адрес родительского семейства веб-сайтов:</span><span class="sxs-lookup"><span data-stu-id="b8642-151">Specify just the URL of the parent site collection:</span></span>
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. <span data-ttu-id="b8642-p107">Закройте приложение Office и снова запустите его. Каталог надстроек будет доступен в диалоговом окне **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="b8642-p107">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="b8642-154">Кроме того, администратор может указать каталог надстроек Office в SharePoint с помощью групповой политики.</span><span class="sxs-lookup"><span data-stu-id="b8642-154">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="b8642-155">Дополнительные сведения см. в разделе [Использование групповой политики для управления возможностью установки и использования пользователями надстроек для Office](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span><span class="sxs-lookup"><span data-stu-id="b8642-155">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office) on TechNet.</span></span>
