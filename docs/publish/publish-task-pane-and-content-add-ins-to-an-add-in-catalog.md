---
title: Публикация надстроек области задач и контентных надстроек в каталоге SharePoint
description: Чтобы делать надстройки Office доступными пользователям в организации, администраторы могут отправлять файлы манифестов надстроек Office в соответствующий каталог надстроек.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: ed4f9778e4cd7dccba00d2e8c019bd4441b70eeb
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870963"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="2720a-103">Публикация надстроек области задач и контентных надстроек в каталоге SharePoint</span><span class="sxs-lookup"><span data-stu-id="2720a-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="2720a-p101">Каталог надстроек — это отдельное семейство веб-сайтов в веб-приложении SharePoint или клиенте SharePoint Online, в котором размещены библиотеки документов для надстроек Office и SharePoint. Администраторы могут отправлять в него файлы манифестов надстроек Office, чтобы пользователи в организации могли получить доступ к этим надстройкам. Когда администратор зарегистрирует каталог надстроек как доверенный, пользователи смогут вставлять надстройки в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="2720a-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="2720a-106">Каталоги надстроек в SharePoint не поддерживают функции надстроек, реализованные в узле `VersionOverrides` [манифеста надстройки](../develop/add-in-manifests.md), такие как команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="2720a-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="2720a-107">Чтобы публиковать надстройки для облачной или гибридной среды, рекомендуем использовать [централизованное развертывание через Центр администрирования Office 365](../publish/centralized-deployment.md).</span><span class="sxs-lookup"><span data-stu-id="2720a-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="2720a-108">Каталоги SharePoint не поддерживаются в Office для Mac.</span><span class="sxs-lookup"><span data-stu-id="2720a-108">SharePoint catalogs are not supported for Office for Mac.</span></span> <span data-ttu-id="2720a-109">Для развертывания надстроек Office на клиентах Mac необходимо отправить их в [AppSource](/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="2720a-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="set-up-an-add-in-catalog"></a><span data-ttu-id="2720a-110">Настройка каталога надстроек</span><span class="sxs-lookup"><span data-stu-id="2720a-110">Set up an add-in catalog</span></span>

<span data-ttu-id="2720a-111">Выполните действия, описанные в одном из указанных ниже разделов, чтобы настроить каталог надстроек в SharePoint или Office 365.</span><span class="sxs-lookup"><span data-stu-id="2720a-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-set-up-an-add-in-catalog-for-on-premises-sharepoint"></a><span data-ttu-id="2720a-112">Настройка каталога надстроек в локальном SharePoint</span><span class="sxs-lookup"><span data-stu-id="2720a-112">To set up an add-in catalog for on-premises SharePoint</span></span>

> [!NOTE]
> <span data-ttu-id="2720a-113">Надстройки в пользовательском интерфейсе локального SharePoint по-прежнему называются **приложениями**.</span><span class="sxs-lookup"><span data-stu-id="2720a-113">The UI in on-premises SharePoint still refers to add-ins as **apps**.</span></span>

1. <span data-ttu-id="2720a-114">Перейдите на **сайт центра администрирования**.</span><span class="sxs-lookup"><span data-stu-id="2720a-114">Browse to the  **Central Administration Site**.</span></span>

2. <span data-ttu-id="2720a-115">В области задач слева выберите пункт **Приложения**.</span><span class="sxs-lookup"><span data-stu-id="2720a-115">In the left task pane, choose  **Apps**.</span></span>

3. <span data-ttu-id="2720a-116">На странице **Приложения** в разделе **Управление приложениями** выберите пункт **Управление каталогом приложений**.</span><span class="sxs-lookup"><span data-stu-id="2720a-116">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>

4. <span data-ttu-id="2720a-117">На странице  **Управление каталогом приложений** убедитесь, что в пункте **Селектор веб-приложения** выбрано правильное веб-приложение.</span><span class="sxs-lookup"><span data-stu-id="2720a-117">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>

5. <span data-ttu-id="2720a-118">Выберите элемент **Просмотреть параметры сайта**.</span><span class="sxs-lookup"><span data-stu-id="2720a-118">Choose  **View site settings**.</span></span>

6. <span data-ttu-id="2720a-119">На странице  **Параметры сайта** выберите пункт **Администраторы семейства веб-сайтов**, чтобы указать администраторов семейства веб-сайтов, а затем нажмите кнопку  **ОК**.</span><span class="sxs-lookup"><span data-stu-id="2720a-119">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>

7. <span data-ttu-id="2720a-120">Чтобы предоставить пользователям разрешения для сайтов, последовательно выберите элементы  **Разрешения для сайта** и **Предоставить разрешения**.</span><span class="sxs-lookup"><span data-stu-id="2720a-120">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>

8. <span data-ttu-id="2720a-121">В диалоговом окне  **Общий доступ к сайту каталога приложений** укажите одного или нескольких пользователей сайта, задайте соответствующие разрешения для них, при необходимости укажите другие параметры, а затем выберите элемент **Общий доступ**.</span><span class="sxs-lookup"><span data-stu-id="2720a-121">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>

9. <span data-ttu-id="2720a-122">Чтобы добавить надстройку в каталог надстроек Office, выберите **Приложения для Office**.</span><span class="sxs-lookup"><span data-stu-id="2720a-122">To add an add-in to the Office Add-ins add-in catalog, choose **Apps for Office**.</span></span>

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a><span data-ttu-id="2720a-123">Настройка каталога надстроек в Office 365</span><span class="sxs-lookup"><span data-stu-id="2720a-123">To set up an add-in catalog on Office 365</span></span>

1. <span data-ttu-id="2720a-124">На странице Центра администрирования Office 365 выберите элемент **Администратор**, а затем **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="2720a-124">On the Office 365 admin center page, choose  **Admin**, and then choose  **SharePoint**.</span></span>

2. <span data-ttu-id="2720a-125">В области задач слева выберите пункт  **надстройки**.</span><span class="sxs-lookup"><span data-stu-id="2720a-125">In the left task pane, choose  **add-ins**.</span></span>

3. <span data-ttu-id="2720a-126">На странице  **надстройки** выберите пункт **Каталог надстроек**.</span><span class="sxs-lookup"><span data-stu-id="2720a-126">On the  **add-ins** page, choose **Add-in Catalog**.</span></span>

4. <span data-ttu-id="2720a-127">На странице  **Сайт каталога надстроек** нажмите кнопку **ОК**, чтобы принять параметр по умолчанию и создать сайт каталога надстроек.</span><span class="sxs-lookup"><span data-stu-id="2720a-127">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>

5. <span data-ttu-id="2720a-128">На странице  **Создание семейства веб-сайтов каталога надстроек** укажите название сайта каталога надстроек.</span><span class="sxs-lookup"><span data-stu-id="2720a-128">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

6. <span data-ttu-id="2720a-129">Укажите адрес веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="2720a-129">Specify the web site address.</span></span>

7. <span data-ttu-id="2720a-p103">Минимальное допустимое значение (в данный момент оно составляет 110) указано в параметре  **Дисковая квота**. В этом семействе веб-сайтов будут устанавливаться только пакеты надстройка, которые имеют небольшой размер.</span><span class="sxs-lookup"><span data-stu-id="2720a-p103">Set the  **Storage Quota** to the lowest possible value (currently 110). You will only be installing add-in packages on this site collection and they are very small.</span></span>

8. <span data-ttu-id="2720a-p104">Задайте для параметра  **Квота ресурсов сервера** значение 0 (ноль). (Квота ресурсов сервера связана с регулированием изолированных решений с низкой производительностью, но на сайте каталога надстроек не будут устанавливаться изолированные решения.)</span><span class="sxs-lookup"><span data-stu-id="2720a-p104">Set the  **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

9. <span data-ttu-id="2720a-134">Нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="2720a-134">Choose  **OK**.</span></span>

10. <span data-ttu-id="2720a-p105">Чтобы добавить надстройку на сайт каталога надстроек, перейдите на только что созданный сайт. В области навигации слева выберите пункт **Надстройки для Office**, а затем выберите команду **новая надстройка**, чтобы отправить надстройку для файла манифеста Office.</span><span class="sxs-lookup"><span data-stu-id="2720a-p105">To add an add-in to the Add-in Catalog Site, browse to the site you have just created. In the left navigation pane, choose  **Office Add-ins**, and then, to upload an Office Add-in manifest file, choose  **new add-in**.</span></span>

## <a name="publish-an-add-in-to-an-add-in-catalog"></a><span data-ttu-id="2720a-137">Публикация надстройки в каталоге надстроек</span><span class="sxs-lookup"><span data-stu-id="2720a-137">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="2720a-138">Чтобы опубликовать надстройку в каталоге надстроек, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="2720a-138">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="2720a-139">Перейдите в каталог надстроек.</span><span class="sxs-lookup"><span data-stu-id="2720a-139">Browse to the add-in catalog:</span></span>

    - <span data-ttu-id="2720a-140">Откройте главную страницу центра администрирования SharePoint.</span><span class="sxs-lookup"><span data-stu-id="2720a-140">Open the SharePoint Central Administration main page.</span></span>

    - <span data-ttu-id="2720a-141">Выберите **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="2720a-141">Select  **Add-ins**.</span></span>

    - <span data-ttu-id="2720a-142">Выберите **Управление каталогом надстроек**.</span><span class="sxs-lookup"><span data-stu-id="2720a-142">Select  **Manage Add-in Catalog**.</span></span>

    - <span data-ttu-id="2720a-143">Выберите указанную ссылку, а затем нажмите **Надстройки Office** на левой панели навигации.</span><span class="sxs-lookup"><span data-stu-id="2720a-143">Choose the link provided, and then choose  **Office Add-ins** on the left navigation bar.</span></span>

2. <span data-ttu-id="2720a-144">Выберите ссылку **Щелкните для добавления нового элемента**.</span><span class="sxs-lookup"><span data-stu-id="2720a-144">Choose the  **Click to add new item** link.</span></span>

3. <span data-ttu-id="2720a-145">Нажмите кнопку **Обзор**, а затем укажите [манифест](../develop/add-in-manifests.md) для отправки.</span><span class="sxs-lookup"><span data-stu-id="2720a-145">Choose  **Browse**, and then specify the [manifest](../develop/add-in-manifests.md) to upload.</span></span>

    <span data-ttu-id="2720a-p106">Теперь надстройки области задач и контентные надстройки из этого каталога доступны в диалоговом окне **Надстройки Office**. Для доступа к ним выберите**Мои надстройки** на вкладке **Вставка**, а затем нажмите **Моя организация**.</span><span class="sxs-lookup"><span data-stu-id="2720a-p106">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="2720a-148">Работа пользователей с каталогом надстроек</span><span class="sxs-lookup"><span data-stu-id="2720a-148">End user experience with the add-in catalog</span></span>

<span data-ttu-id="2720a-149">Пользователь может получить доступ к каталогу надстроек в приложении Office, выполнив указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="2720a-149">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="2720a-150">В приложении Office выберите **Файл** > **Параметры** > **Центр управления безопасностью** > **Параметры центра управления безопасностью** > **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="2720a-150">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>

2. <span data-ttu-id="2720a-151">Укажите URL-адрес _родительского семейства веб-сайтов SharePoint_ для каталога надстроек.</span><span class="sxs-lookup"><span data-stu-id="2720a-151">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 

    <span data-ttu-id="2720a-152">Предположим, что URL-адрес каталога надстроек Office такой:</span><span class="sxs-lookup"><span data-stu-id="2720a-152">For example, if the URL of the Office Add-ins catalog is:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`

    <span data-ttu-id="2720a-153">Укажите только URL-адрес родительского семейства веб-сайтов:</span><span class="sxs-lookup"><span data-stu-id="2720a-153">Specify just the URL of the parent site collection:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`

3. <span data-ttu-id="2720a-p107">Закройте приложение Office и снова запустите его. Каталог надстроек будет доступен в диалоговом окне **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="2720a-p107">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="2720a-156">Кроме того, администратор может указать каталог надстроек Office в SharePoint с помощью групповой политики.</span><span class="sxs-lookup"><span data-stu-id="2720a-156">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="2720a-157">Дополнительные сведения см. в разделе [Использование групповой политики для управления возможностью установки и использования пользователями приложений для Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span><span class="sxs-lookup"><span data-stu-id="2720a-157">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span></span>
