---
title: Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET
description: ''
ms.date: 12/04/2019
localization_priority: Priority
ms.openlocfilehash: 6306616880138574ede8a127b7fd3c2a65061191
ms.sourcegitcommit: 43166612e9b4bf7a73312a572663c8696353dbc6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/29/2020
ms.locfileid: "41580982"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="c2412-102">Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="c2412-102">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="c2412-103">После того как пользователи войдут в Office, ваша надстройка сможет использовать те же учетные данные для предоставления им доступа к нескольким приложениям без необходимости повторного входа.</span><span class="sxs-lookup"><span data-stu-id="c2412-103">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time.</span></span> <span data-ttu-id="c2412-104">Общие сведения см. в статье [Включение единого входа в надстройке Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="c2412-104">For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>
<span data-ttu-id="c2412-105">Из этой статьи вы узнаете, как включить единый вход в надстройке, созданной с помощью Node.js и Express.</span><span class="sxs-lookup"><span data-stu-id="c2412-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span>

> [!NOTE]
> <span data-ttu-id="c2412-106">Аналогичная статья, посвященная надстройке на основе ASP.NET, — [Создание надстройки Office на платформе Node.js с использованием единого входа](create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="c2412-106">For a similar article about an ASP.NET-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c2412-107">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="c2412-107">Prerequisites</span></span>

* <span data-ttu-id="c2412-108">Visual Studio 2019 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="c2412-108">Visual Studio 2019 or later.</span></span>

* [<span data-ttu-id="c2412-109">Office Developer Tools</span><span class="sxs-lookup"><span data-stu-id="c2412-109">Office Developer Tools</span></span>](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="c2412-110">Несколько файлов и папок, сохраненных в OneDrive для бизнеса в составе подписки на Office 365.</span><span class="sxs-lookup"><span data-stu-id="c2412-110">At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.</span></span>

* <span data-ttu-id="c2412-111">Подписка на Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="c2412-111">A Microsoft Azure subscription.</span></span> <span data-ttu-id="c2412-112">Эта надстройка требует наличия Azure Active Directory (AD).</span><span class="sxs-lookup"><span data-stu-id="c2412-112">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="c2412-113">В Azure AD доступны службы идентификации, которые приложения используют для проверки подлинности и авторизации.</span><span class="sxs-lookup"><span data-stu-id="c2412-113">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="c2412-114">Пробную подписку можно получить на сайте [Microsoft Azure](https://account.windowsazure.com/SignUp).</span><span class="sxs-lookup"><span data-stu-id="c2412-114">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="c2412-115">Настройка начального проекта</span><span class="sxs-lookup"><span data-stu-id="c2412-115">Set up the starter project</span></span>

<span data-ttu-id="c2412-116">Клонируйте или скачайте репозиторий [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span><span class="sxs-lookup"><span data-stu-id="c2412-116">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

> [!NOTE]
> <span data-ttu-id="c2412-117">Существует две версии примера.</span><span class="sxs-lookup"><span data-stu-id="c2412-117">There are two versions of the sample:</span></span>
>
> * <span data-ttu-id="c2412-p103">В папке **Before** находится начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы. В последующих разделах этой статьи рассматривается доработка проекта.</span><span class="sxs-lookup"><span data-stu-id="c2412-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
> * <span data-ttu-id="c2412-121">Версия примера в папке **Complete** идентична надстройке, которую вы бы создали, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью.</span><span class="sxs-lookup"><span data-stu-id="c2412-121">The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="c2412-122">Чтобы использовать готовую версию, просто выполните действия, описанные в этой статье, но замените папку "Before" на папку "Complete" и пропустите разделы **Код на стороне клиента** и **Код на стороне сервера**.</span><span class="sxs-lookup"><span data-stu-id="c2412-122">To use the completed version, just follow the instructions in this article, but replace "Before" with "Complete" and skip the sections **Code the client side** and **Code the server side**.</span></span>


## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="c2412-123">Регистрация надстройки в конечной точке Azure AD версии 2.0</span><span class="sxs-lookup"><span data-stu-id="c2412-123">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="c2412-124">Перейдите на страницу [регистрации приложений портала Azure](https://go.microsoft.com/fwlink/?linkid=2083908), чтобы зарегистрировать свое приложение.</span><span class="sxs-lookup"><span data-stu-id="c2412-124">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="c2412-125">Войдите в клиент Office 365, используя учетные данные ***администратора***.</span><span class="sxs-lookup"><span data-stu-id="c2412-125">Sign in with the ***admin*** credentials to your Office 365 tenancy.</span></span> <span data-ttu-id="c2412-126">Пример: MyName@contoso.onmicrosoft.com.</span><span class="sxs-lookup"><span data-stu-id="c2412-126">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="c2412-127">Выберите **Новая регистрация**.</span><span class="sxs-lookup"><span data-stu-id="c2412-127">Select **New registration**.</span></span> <span data-ttu-id="c2412-128">На странице**Зарегистрировать приложение** задайте необходимые значения следующим образом.</span><span class="sxs-lookup"><span data-stu-id="c2412-128">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="c2412-129">Введите **имя** `Office-Add-in-ASPNET-SSO`.</span><span class="sxs-lookup"><span data-stu-id="c2412-129">Set **Name** to `Office-Add-in-ASPNET-SSO`.</span></span>
    * <span data-ttu-id="c2412-130">Для параметра **Поддерживаемые типы учетных записей** укажите вариант **Учетные записи в любом каталоге организации (любой каталог Azure AD — мультитенантный) и личные учетные записи Майкрософт (например, Skype, Xbox)**.</span><span class="sxs-lookup"><span data-stu-id="c2412-130">Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.</span></span> <span data-ttu-id="c2412-131">(Если вы хотите, чтобы надстройка была доступна пользователям только в клиенте, в котором вы ее регистрируете, можно выбрать вариант **Учетные записи только в этом каталоге организации…**, но вам потребуется выполнить дополнительные действия по настройке.</span><span class="sxs-lookup"><span data-stu-id="c2412-131">(If you want the add-in to be usable only by users in the tenancy where you are registering it, you can choose **Accounts in this organizational directory only ...** instead, but you will need to go through some additional setup steps.</span></span> <span data-ttu-id="c2412-132">См. раздел **Настройка в однотенантном режиме** ниже.)</span><span class="sxs-lookup"><span data-stu-id="c2412-132">See **Setup for single-tenant** below.)</span></span>
    * <span data-ttu-id="c2412-133">Убедитесь, что в разделе **URI перенаправления** в раскрывающемся списке выбран пункт **Интернет**, и задайте для URI значение ` https://localhost:44355/AzureADAuth/Authorize`.</span><span class="sxs-lookup"><span data-stu-id="c2412-133">In the **Redirect URI** section, ensure that **Web** is selected in the drop down and then set the URI to` https://localhost:44355/AzureADAuth/Authorize`.</span></span>
    * <span data-ttu-id="c2412-134">Нажмите кнопку **Зарегистрировать**.</span><span class="sxs-lookup"><span data-stu-id="c2412-134">Choose **Register**.</span></span>

1. <span data-ttu-id="c2412-135">На странице **Office-Add-in-NodeJS-SSO** скопируйте и сохраните значения параметров **Идентификатор приложения (клиент)** и **Идентификатор каталога (клиент)**.</span><span class="sxs-lookup"><span data-stu-id="c2412-135">On the **Office-Add-in-NodeJS-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="c2412-136">Они понадобятся вам позже.</span><span class="sxs-lookup"><span data-stu-id="c2412-136">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="c2412-137">Этот идентификатор представляет собой значение аудитории, используемое, когда другие приложения, например ведущее приложение Office (PowerPoint, Word, Excel и т. д.), пытаются получить авторизованный доступ к вашему приложению.</span><span class="sxs-lookup"><span data-stu-id="c2412-137">This ID is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="c2412-138">Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="c2412-138">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="c2412-139">В разделе **Управление** выберите **Сертификаты и секреты**.</span><span class="sxs-lookup"><span data-stu-id="c2412-139">Under **Manage**, select **Certificates & secrets**.</span></span> <span data-ttu-id="c2412-140">Нажмите кнопку **Новый секрет клиента**.</span><span class="sxs-lookup"><span data-stu-id="c2412-140">Select the **New client secret** button.</span></span> <span data-ttu-id="c2412-141">Введите значение параметра **Описание**, выберите соответствующий вариант для параметра **Истекает срок действия** и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="c2412-141">Enter a value for **Description**, then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="c2412-142">*Сразу скопируйте значение секрета клиента и сохраните его с идентификатором приложения* перед продолжением, так как он понадобится вам позже.</span><span class="sxs-lookup"><span data-stu-id="c2412-142">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="c2412-143">В разделе **Управление** выберите **Предоставление API**.</span><span class="sxs-lookup"><span data-stu-id="c2412-143">Under **Manage**, select **Expose an API**.</span></span> <span data-ttu-id="c2412-144">Щелкните ссылку **Задать**, чтобы создать URI идентификатора приложения в формате "api://$ИД приложения GUID$", где $App ID GUID$ — **идентификатор приложения (клиента)**.</span><span class="sxs-lookup"><span data-stu-id="c2412-144">Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span> <span data-ttu-id="c2412-145">Вставьте `localhost:44355/` (обратите внимание на знак косой черты "/", добавленный в конце) после `//` и перед GUID.</span><span class="sxs-lookup"><span data-stu-id="c2412-145">Insert `localhost:44355/` (note the forward slash "/" appended to the end) after the `//` and before the GUID.</span></span> <span data-ttu-id="c2412-146">Весь идентификатор должен отображаться в формате `api://localhost:44355/$App ID GUID$`, например: `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span><span class="sxs-lookup"><span data-stu-id="c2412-146">The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

1. <span data-ttu-id="c2412-147">В диалоговом окне выберите **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="c2412-147">Select **Save** on the dialog.</span></span>

1. <span data-ttu-id="c2412-148">Нажмите кнопку **Добавить область**.</span><span class="sxs-lookup"><span data-stu-id="c2412-148">Select the **Add a scope** button.</span></span> <span data-ttu-id="c2412-149">В открывшейся панели введите `access_as_user` в качестве параметра **Имя области**.</span><span class="sxs-lookup"><span data-stu-id="c2412-149">In the panel that opens, enter `access_as_user` as the **Scope** name.</span></span>

1. <span data-ttu-id="c2412-150">Для параметра **Кто может давать согласие?** установите вариант **Администраторы и пользователи**.</span><span class="sxs-lookup"><span data-stu-id="c2412-150">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="c2412-151">Заполните поля для настройки запросов согласия администраторов и пользователей значениями, соответствующими области `access_as_user`, позволяющей ведущему приложению Office использовать веб-интерфейсы API надстройки с такими же правами, как у текущего пользователя.</span><span class="sxs-lookup"><span data-stu-id="c2412-151">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office host application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="c2412-152">Возможные варианты:</span><span class="sxs-lookup"><span data-stu-id="c2412-152">Suggestions:</span></span>

    - <span data-ttu-id="c2412-153">**Отображаемое имя согласия администратора**. Office может действовать в качестве пользователя.</span><span class="sxs-lookup"><span data-stu-id="c2412-153">**Admin consent title**: Office can act as the user.</span></span>
    - <span data-ttu-id="c2412-154">**Описание согласия администратора**. Позволяет Office вызывать веб-API надстройки с такими же правами, как у текущего пользователя.</span><span class="sxs-lookup"><span data-stu-id="c2412-154">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    - <span data-ttu-id="c2412-155">**Отображаемое имя согласия пользователя**. Office может действовать от вашего имени.</span><span class="sxs-lookup"><span data-stu-id="c2412-155">**User consent title**: Office can act as you.</span></span>
    - <span data-ttu-id="c2412-156">**Описание согласия администратора**. Позволяет Office вызывать веб-API надстройки с такими же правами, как у вас.</span><span class="sxs-lookup"><span data-stu-id="c2412-156">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="c2412-157">Убедитесь, что параметру **Состояние** присвоено значение **Включено**.</span><span class="sxs-lookup"><span data-stu-id="c2412-157">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="c2412-158">Нажмите кнопку **Добавить область**.</span><span class="sxs-lookup"><span data-stu-id="c2412-158">Select **Add scope** .</span></span>

    > [!NOTE]
    > <span data-ttu-id="c2412-159">Доменная часть имени **области**, отображаемая непосредственно под текстовым полем, должна автоматически соответствовать URI идентификатора приложения, заданного ранее, с добавлением `/access_as_user` в конце, например: `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="c2412-159">The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="c2412-160">В разделе **Авторизованные клиентские приложения** укажите приложения, которые необходимо авторизовать для веб-приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2412-160">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="c2412-161">Необходимо обеспечить предварительную авторизацию для всех указанных ниже идентификаторов.</span><span class="sxs-lookup"><span data-stu-id="c2412-161">Each of the following IDs needs to be pre-authorized.</span></span>

    - <span data-ttu-id="c2412-162">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="c2412-162">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="c2412-163">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="c2412-163">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="c2412-164">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="c2412-164">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="c2412-165">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook в Интернете).</span><span class="sxs-lookup"><span data-stu-id="c2412-165">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span></span>

    <span data-ttu-id="c2412-166">Для каждого идентификатора сделайте следующее:</span><span class="sxs-lookup"><span data-stu-id="c2412-166">For each ID, take these steps:</span></span>

    <span data-ttu-id="c2412-167">а)</span><span class="sxs-lookup"><span data-stu-id="c2412-167">a.</span></span> <span data-ttu-id="c2412-168">Нажмите кнопку **Добавить клиентское приложение**, в открывшейся панели присвойте параметру "Идентификатор клиента" соответствующий код GUID и установите флажок `api://localhost:44355/$App ID GUID$/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="c2412-168">Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="c2412-169">б)</span><span class="sxs-lookup"><span data-stu-id="c2412-169">b.</span></span> <span data-ttu-id="c2412-170">Нажмите кнопку **Добавить приложение**.</span><span class="sxs-lookup"><span data-stu-id="c2412-170">Select **Add application**.</span></span>

1. <span data-ttu-id="c2412-171">В разделе **Управление** выберите **Разрешения API** и нажмите кнопку **Добавить разрешение**.</span><span class="sxs-lookup"><span data-stu-id="c2412-171">Under **Manage**, select **API permissions** and then select **Add a permission**.</span></span> <span data-ttu-id="c2412-172">В открывшейся панели выберите **Microsoft Graph** и щелкните **Делегированные разрешения**.</span><span class="sxs-lookup"><span data-stu-id="c2412-172">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="c2412-173">Используйте поле поиска **Выбрать разрешения**, чтобы найти нужные разрешения для надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2412-173">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="c2412-174">Выберите следующие параметры.</span><span class="sxs-lookup"><span data-stu-id="c2412-174">Select the following.</span></span> <span data-ttu-id="c2412-175">Для самой надстройки требуется только первое разрешение, но разрешение `profile` необходимо, чтобы ведущее приложение Office получило маркер для веб-приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2412-175">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span> <span data-ttu-id="c2412-176">(Для надстройки требуются только разрешения Files.Read.All и profile.</span><span class="sxs-lookup"><span data-stu-id="c2412-176">(Only Files.Read.All and profile are actually needed by the add-in.</span></span> <span data-ttu-id="c2412-177">Остальные два необходимо запросить для библиотеки MSAL.NET.)</span><span class="sxs-lookup"><span data-stu-id="c2412-177">You must request the other two because the MSAL.NET library requires them.)</span></span>

    * <span data-ttu-id="c2412-178">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="c2412-178">Files.Read.All</span></span>
    * <span data-ttu-id="c2412-179">offline_access</span><span class="sxs-lookup"><span data-stu-id="c2412-179">offline_access</span></span>
    * <span data-ttu-id="c2412-180">openid</span><span class="sxs-lookup"><span data-stu-id="c2412-180">openid</span></span>
    * <span data-ttu-id="c2412-181">profile</span><span class="sxs-lookup"><span data-stu-id="c2412-181">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="c2412-182">Разрешение `User.Read` может быть уже указано по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="c2412-182">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="c2412-183">Незачем запрашивать ненужные разрешения, поэтому рекомендуем снять флажок рядом с разрешением, которое не требуется вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="c2412-183">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="c2412-184">Установите флажок для каждого отображаемого разрешения.</span><span class="sxs-lookup"><span data-stu-id="c2412-184">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="c2412-185">Выбрав нужные для надстройки разрешения, нажмите кнопку **Добавить разрешения** в нижней части панели.</span><span class="sxs-lookup"><span data-stu-id="c2412-185">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="c2412-186">На этой же странице нажмите кнопку **Предоставить согласие администратора для [имя клиента]** и выберите **Принять** в появившемся запросе подтверждения.</span><span class="sxs-lookup"><span data-stu-id="c2412-186">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Accept** for the confirmation that appears.</span></span>

    > [!NOTE]
    > <span data-ttu-id="c2412-187">После нажатия кнопки **Предоставить согласие администратора для [имя клиента]** может появиться сообщение баннера с просьбой повторить попытку через несколько минут, чтобы можно было создать запрос на продолжение.</span><span class="sxs-lookup"><span data-stu-id="c2412-187">After choosing **Grant admin consent for [tenant name]**, you may see a banner message asking you to try again in a few minutes so that the consent prompt can be constructed.</span></span> <span data-ttu-id="c2412-188">В этом случае вы можете перейти к следующему разделу, ***но не забудьте вернуться на портал и нажать эту кнопку***!</span><span class="sxs-lookup"><span data-stu-id="c2412-188">If so, you can start work on the next section, ***but don't forget to come back to the portal and press this button***!</span></span>

## <a name="configure-the-solution"></a><span data-ttu-id="c2412-189">Настройка решения</span><span class="sxs-lookup"><span data-stu-id="c2412-189">Configure the solution</span></span>

1. <span data-ttu-id="c2412-190">В корне папки **Before** откройте SLN-файл решения в **Visual Studio**.</span><span class="sxs-lookup"><span data-stu-id="c2412-190">In the root of the **Before** folder, open the solution (.sln) file in **Visual Studio**.</span></span> <span data-ttu-id="c2412-191">В **обозревателе решений** щелкните правой кнопкой мыши верхний узел (узел решения, а не узлы проектов) и выберите **Назначить запускаемые проекты**.</span><span class="sxs-lookup"><span data-stu-id="c2412-191">Right-click the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.</span></span>

1. <span data-ttu-id="c2412-192">В разделе **Общие свойства** выберите **Запускаемый проект**, а затем **Несколько запускаемых проектов**.</span><span class="sxs-lookup"><span data-stu-id="c2412-192">Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**.</span></span> <span data-ttu-id="c2412-193">Убедитесь, что для параметра **Действие** в обоих проектах установлено значение **Запуск** и что проект, заканчивающийся на "...WebAPI", указан в списке первым.</span><span class="sxs-lookup"><span data-stu-id="c2412-193">Ensure that the **Action** for both projects is set to **Start**, and that the project that ends in "...WebAPI" is listed first.</span></span> <span data-ttu-id="c2412-194">Закройте диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="c2412-194">Close the dialog.</span></span>

1. <span data-ttu-id="c2412-195">Вернувшись в **Обозреватель решений**, выберите (не используя правую кнопку мыши) проект **Office-Add-in-Microsoft-Graph-ASPNETWebAPI**.</span><span class="sxs-lookup"><span data-stu-id="c2412-195">Back in **Solution Explorer**, select (don't right-click) the **Office-Add-in-Microsoft-Graph-ASPNETWebAPI** project.</span></span> <span data-ttu-id="c2412-196">Откроется область **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="c2412-196">The **Properties** pane opens.</span></span> <span data-ttu-id="c2412-197">Убедитесь, что для параметра **SSL включен** задано значение **True**.</span><span class="sxs-lookup"><span data-stu-id="c2412-197">Ensure that **SSL Enabled** is **True**.</span></span> <span data-ttu-id="c2412-198">Убедитесь, что **URL-адрес SSL** указан как `http://localhost:44355/`.</span><span class="sxs-lookup"><span data-stu-id="c2412-198">Verify that the **SSL URL** is `http://localhost:44355/`.</span></span>

1. <span data-ttu-id="c2412-199">В файле web.config используйте значения, скопированные ранее.</span><span class="sxs-lookup"><span data-stu-id="c2412-199">In "Web.config", use the values that you copied in earlier.</span></span> <span data-ttu-id="c2412-200">Для **ida:ClientID** и **ida:Audience** укажите **идентификатор приложения (клиента)**, для **ida:Password** — секрет клиента.</span><span class="sxs-lookup"><span data-stu-id="c2412-200">Set both the **ida:ClientID** and the **ida:Audience** to your **Application (client) ID**, and set **ida:Password** to your client secret.</span></span>

    > [!NOTE]
    > <span data-ttu-id="c2412-201">**Идентификатор приложения (клиента)** представляет собой значение аудитории, используемое, когда другие приложения, например ведущее приложение Office (PowerPoint, Word, Excel), пытаются получить авторизованный доступ к вашему приложению.</span><span class="sxs-lookup"><span data-stu-id="c2412-201">The **Application (client) ID** is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="c2412-202">Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="c2412-202">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="c2412-203">Если вы не указали вариант "Учетные записи только в этом каталоге организации" для параметра **ПОДДЕРЖИВАЕМЫЕ ТИПЫ УЧЕТНЫХ ЗАПИСЕЙ** при регистрации настройки, сохраните и закройте файл web.config. В противном случае сохраните его, но оставьте открытым. </span><span class="sxs-lookup"><span data-stu-id="c2412-203">If you didn't choose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, save and close the web.config. Otherwise, save but leave it open.</span></span>

1. <span data-ttu-id="c2412-204">В **обозревателе решений** выберите проект **Office-Add-in-Microsoft-Graph-ASPNET** и откройте файл манифеста надстройки Office-Add-in-ASPNET-SSO.xml, а затем прокрутите вниз до конца файла. </span><span class="sxs-lookup"><span data-stu-id="c2412-204">Still in **Solution Explorer**, choose the **Office-Add-in-Microsoft-Graph-ASPNET** project and open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file.</span></span> <span data-ttu-id="c2412-205">Над закрывающим тегом `</VersionOverrides>` вы найдете следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="c2412-205">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="c2412-206">Замените заполнитель "$application_GUID here$" *в обоих местах* разметки идентификатором приложения, скопированным при регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2412-206">Replace the placeholder “$application_GUID here$” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="c2412-207">Символы "$" не входят в состав идентификатора, их не нужно вставлять.</span><span class="sxs-lookup"><span data-stu-id="c2412-207">The "$" signs are not part of the ID, so do not include them.</span></span> <span data-ttu-id="c2412-208">Это тот же идентификатор, который использовался для ClientID и Audience в файле web.config.</span><span class="sxs-lookup"><span data-stu-id="c2412-208">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

  > [!NOTE]
  > <span data-ttu-id="c2412-209">Значение **Resource** — это **URI идентификатора приложения**, указанный при регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2412-209">The **Resource** value is the **Application ID URI** you set when you registered the add-in.</span></span> <span data-ttu-id="c2412-210">Раздел **Scopes** используется для создания диалогового окна согласия, только если надстройка продается в AppSource.</span><span class="sxs-lookup"><span data-stu-id="c2412-210">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="c2412-211">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="c2412-211">Save and close the file.</span></span>

### <a name="setup-for-single-tenant"></a><span data-ttu-id="c2412-212">Настройка в однотенантном режиме</span><span class="sxs-lookup"><span data-stu-id="c2412-212">Setup for single-tenant</span></span>

<span data-ttu-id="c2412-213">Если вы указали вариант "Учетные записи только в этом каталоге организации" для параметра **ПОДДЕРЖИВАЕМЫЕ ТИПЫ УЧЕТНЫХ ЗАПИСЕЙ** при регистрации надстройки, необходимо выполнить дополнительные шаги настройки. </span><span class="sxs-lookup"><span data-stu-id="c2412-213">If you chose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, you need to take these additional setup steps:</span></span>

1. <span data-ttu-id="c2412-214">Вернитесь на портал Azure и откройте колонку **Обзор** регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2412-214">Go back to the Azure Portal and open the **Overview** blade of the add-in's registration.</span></span> <span data-ttu-id="c2412-215">Скопируйте **Идентификатор каталога (клиента)**.</span><span class="sxs-lookup"><span data-stu-id="c2412-215">Copy the **Directory (tenant) ID**.</span></span>

1. <span data-ttu-id="c2412-216">В файле web.config замените "common" в значении **ida:Authority** на GUID, скопированный на предыдущем шаге.  </span><span class="sxs-lookup"><span data-stu-id="c2412-216">In the web.config, replace the "common" in the value of **ida:Authority** with the GUID you copied in the preceding step.</span></span> <span data-ttu-id="c2412-217">После этого значение должно выглядеть следующим образом: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.</span><span class="sxs-lookup"><span data-stu-id="c2412-217">When you are finished the value should look similar to this: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.</span></span>

1. <span data-ttu-id="c2412-218">Сохраните и закройте файл web.config.</span><span class="sxs-lookup"><span data-stu-id="c2412-218">Save and close the web.config.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="c2412-219">Код на стороне клиента</span><span class="sxs-lookup"><span data-stu-id="c2412-219">Code the client side</span></span>

1. <span data-ttu-id="c2412-220">Откройте файл HomeES6.js в папке **Scripts**.</span><span class="sxs-lookup"><span data-stu-id="c2412-220">Open the HomeES6.js file in the **Scripts** folder.</span></span> <span data-ttu-id="c2412-221">В нем уже есть следующий код:</span><span class="sxs-lookup"><span data-stu-id="c2412-221">It already has some code in it:</span></span>

    * <span data-ttu-id="c2412-222">Полизаполнение, которое назначает объект Office.Promise глобальному объекту window, чтобы надстройка могла работать, если в Office используется пользовательский интерфейс Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="c2412-222">A polyfill that assigns the Office.Promise object to the global window object so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="c2412-223">(Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).)</span><span class="sxs-lookup"><span data-stu-id="c2412-223">(For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).)</span></span>
    * <span data-ttu-id="c2412-224">Назначение методу `Office.initialize`, которое, в свою очередь, назначает обработчик события для нажатия кнопки `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="c2412-224">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="c2412-225">Метод `showResult` для отображения сообщения об ошибке (или данных, возвращаемых из Microsoft Graph) в нижней части области задач.</span><span class="sxs-lookup"><span data-stu-id="c2412-225">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="c2412-226">Метод `logErrors` для регистрации в консоли ошибок, которые не предназначены для пользователя.</span><span class="sxs-lookup"><span data-stu-id="c2412-226">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>
    * <span data-ttu-id="c2412-227">Код для реализации резервной системы авторизации, которая будет использоваться надстройкой в сценариях, где единый вход не поддерживается или возникла ошибка единого входа.</span><span class="sxs-lookup"><span data-stu-id="c2412-227">Code that implements the fallback authorization system that the add-in will use in scenarios where SSO is not supported or has errored.</span></span>

1. <span data-ttu-id="c2412-228">Под назначением методу `Office.initialize` добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="c2412-228">Below the assignment to `Office.initialize`, add the code below.</span></span> <span data-ttu-id="c2412-229">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="c2412-229">Note the following about this code:</span></span>

    * <span data-ttu-id="c2412-230">При обработке ошибок в надстройке иногда автоматически выполняется еще одна попытка получить маркер доступа с помощью другого набора параметров.</span><span class="sxs-lookup"><span data-stu-id="c2412-230">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="c2412-231">Переменная счетчика `retryGetAccessToken` используется, чтобы предотвратить циклическое повторение неудачных попыток получить маркер.</span><span class="sxs-lookup"><span data-stu-id="c2412-231">The counter variable `retryGetAccessToken` is used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="c2412-232">Функция `getGraphData` определяется ключевым словом `async` в ES6.</span><span class="sxs-lookup"><span data-stu-id="c2412-232">The `getGraphData` function is defined with the ES6 `async` keyword.</span></span> <span data-ttu-id="c2412-233">Синтаксис ES6 значительно упрощает использование API единого входа в надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="c2412-233">Using ES6 syntax makes the SSO API in Office Add-ins much easier to to use.</span></span> <span data-ttu-id="c2412-234">Это единственный файл в решении, в котором используется синтаксис, не поддерживаемый в Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="c2412-234">This is the only file in the solution that will use syntax that is not supported by Internet Explorer.</span></span> <span data-ttu-id="c2412-235">"ES6" включается в имя файла в качестве напоминания.</span><span class="sxs-lookup"><span data-stu-id="c2412-235">We put 'ES6' in the filename as a reminder.</span></span> <span data-ttu-id="c2412-236">Компилятор TSC используется в решении для компиляции этого файла в ES5, чтобы надстройка могла работать, если в Office используется пользовательский интерфейс Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="c2412-236">The solution uses the tsc transpiler to transpile this file to ES5, so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="c2412-237">(См. файл tsconfig.json в корневой папке проекта.)</span><span class="sxs-lookup"><span data-stu-id="c2412-237">(See the tsconfig.json file in the root of the project.)</span></span>

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, forMSGraphAccess: true });
    }
    ```

1. <span data-ttu-id="c2412-238">Добавьте указанную ниже функцию под функцией `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="c2412-238">Below the `getGraphData` function add the following function.</span></span> <span data-ttu-id="c2412-239">Обратите внимание, что функция `handleClientSideErrors` будет создана позже.</span><span class="sxs-lookup"><span data-stu-id="c2412-239">Note that you create the `handleClientSideErrors` function in a later step.</span></span>

    ```javascript
    async function getDataWithToken() {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for an access token to Microsoft Graphn and then get the data
            //         from Microsoft Graph.

        }
        catch (exception) {
            if (exception.code) {
                handleClientSideErrors(exception);
            }
            else {
                showResult(["EXCEPTION: " + JSON.stringify(exception)]);
            }
        }
    }
    ```

1. <span data-ttu-id="c2412-240">Замените `TODO 1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-240">Replace `TODO 1` with the following.</span></span> <span data-ttu-id="c2412-241">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="c2412-241">About this code, note:</span></span>

    * <span data-ttu-id="c2412-242">`getAccessToken` предписывает Office получить маркер начальной загрузки из Azure AD и вернуть в надстройку.</span><span class="sxs-lookup"><span data-stu-id="c2412-242">`getAccessToken` tells Office to get a bootstrap token from Azure AD and return to the add-in.</span></span>
    * <span data-ttu-id="c2412-243">`allowSignInPrompt` предписывает Office предложить пользователю выполнить вход, если он еще не вошел в Office.</span><span class="sxs-lookup"><span data-stu-id="c2412-243">`allowSignInPrompt` tells Office to prompt the user to sign in if the user isn't already signed into Office.</span></span>
    * <span data-ttu-id="c2412-244">`forMSGraphAccess` сообщает Office, что надстройка планирует заменить маркер начальной загрузки на маркер доступа к Microsoft Graph (вместо того, чтобы использовать его в качестве маркера ИД пользователя).</span><span class="sxs-lookup"><span data-stu-id="c2412-244">`forMSGraphAccess` tells Office that the add-in intends to swap the bootstrap token for an access token to Microsoft Graph (instead of just using the bootstrap token as a user ID token).</span></span> <span data-ttu-id="c2412-245">Установка этого параметра дает Office возможность отменить процесс получения маркера начальной загрузки (и вернуть код ошибки 13012), если администратор клиента пользователя не предоставил согласие надстройке.</span><span class="sxs-lookup"><span data-stu-id="c2412-245">Setting this option gives Office a chance to cancel the process of getting a bootstrap token (and return error code 13012) if the user's tenant administrator has not granted consent to the add-in.</span></span> <span data-ttu-id="c2412-246">Код на стороне клиента может реагировать на ошибку 13012, переходя на резервную систему авторизации.</span><span class="sxs-lookup"><span data-stu-id="c2412-246">The add-in's client-side code can respond to the 13012 by branching to a fallback authorization system.</span></span> <span data-ttu-id="c2412-247">Если `forMSGraphAccess` не используется и администратор не предоставил согласие, маркер начальной загрузки возвращается, но попытка заменить его в потоке "от имени" может привести к ошибке.</span><span class="sxs-lookup"><span data-stu-id="c2412-247">If the `forMSGraphAccess` is not used, and the admin has not granted consent, the bootstrap token is returned, but the attempt to exhange it with the on-behalf-of flow would result in an error.</span></span> <span data-ttu-id="c2412-248">Таким образом, параметр `forMSGraphAccess` позволяет надстройке быстро перейти на резервную систему.</span><span class="sxs-lookup"><span data-stu-id="c2412-248">Thus, the `forMSGraphAccess` option enables the add-in to branch to the fallback system quickly.</span></span>
    * <span data-ttu-id="c2412-249">Вы создадите функцию `getData` позже.</span><span class="sxs-lookup"><span data-stu-id="c2412-249">You create the `getData` function in a later step.</span></span>
    * <span data-ttu-id="c2412-250">Параметр `/api/values` является URL-адресом контроллера на стороне сервера, который будет осуществлять обмен маркерами и использовать маркер доступа, полученный обратно, для вызова Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="c2412-250">The `/api/values` parameter is the URL of a server-side controller that will make the token exchange and use the access token it gets back to make the call to Microsoft Graph.</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. <span data-ttu-id="c2412-251">Добавьте указанный ниже код под функцией `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="c2412-251">Below the `getGraphData` function, add the following.</span></span> <span data-ttu-id="c2412-252">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="c2412-252">About this code, note:</span></span>

    * <span data-ttu-id="c2412-253">Он используется и в системах единого входа, и в резервных системах авторизации.</span><span class="sxs-lookup"><span data-stu-id="c2412-253">It is used by both the SSO and the fallback authorization systems.</span></span>
    * <span data-ttu-id="c2412-254">Параметр `relativeUrl` является контроллером на стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="c2412-254">The `relativeUrl` parameter is a server-side controller.</span></span>
    * <span data-ttu-id="c2412-255">Параметр `accessToken` может быть маркером начальной загрузки или маркером полного доступа.</span><span class="sxs-lookup"><span data-stu-id="c2412-255">The `accessToken` parameter can be a bootstrap token or a full access token.</span></span>
    * <span data-ttu-id="c2412-256">`writeFileNamesToOfficeDocument` уже включен в проект.</span><span class="sxs-lookup"><span data-stu-id="c2412-256">The `writeFileNamesToOfficeDocument` is already part of the project.</span></span>
    * <span data-ttu-id="c2412-257">Вы создадите функцию `handleServerSideErrors` позже.</span><span class="sxs-lookup"><span data-stu-id="c2412-257">You create the `handleServerSideErrors` function in a later step.</span></span>

    ```javascript
    function getData(relativeUrl, accessToken) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
            .done(function (result) {
                writeFileNamesToOfficeDocument(result)
                    .then(function () {
                        showResult(["Your data has been added to the document."]);
                    })
                    .catch(function (error) {
                        showResult([JSON.stringify(error)]);
                    });
            })
            .fail(function (result) {
                handleServerSideErrors(result);
            });
    }
    ```

### <a name="handle-client-side-errors"></a><span data-ttu-id="c2412-258">Обработка ошибок на стороне клиента</span><span class="sxs-lookup"><span data-stu-id="c2412-258">Handle client-side errors</span></span>

1. <span data-ttu-id="c2412-259">Добавьте указанную ниже функцию под функцией `getData`.</span><span class="sxs-lookup"><span data-stu-id="c2412-259">Below the `getData` function, add the following function.</span></span> <span data-ttu-id="c2412-260">Обратите внимание, что `error.code` — это число (обычно в диапазоне 13xxx).</span><span class="sxs-lookup"><span data-stu-id="c2412-260">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 2: Handle errors where the add-in should NOT invoke
            //         the alternative system of authorization.

            // TODO 3: Handle errors where the add-in should invoke
            //         the alternative system of authorization.

        }
    }
    ```

1. <span data-ttu-id="c2412-261">Замените `TODO 2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-261">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="c2412-262">Дополнительные сведения об этих ошибках см. в статье [Устранение ошибок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="c2412-262">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span>

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
        break;
    case 13006:
        // Only seen in Office on the Web.
        showResult(["Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the Web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the Web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. <span data-ttu-id="c2412-263">Замените `TODO 3` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-263">Replace `TODO 3` with the following code.</span></span> <span data-ttu-id="c2412-264">Во всех других случаях надстройка переходит на резервную систему авторизации.</span><span class="sxs-lookup"><span data-stu-id="c2412-264">For all other errors, the add-in branches to the fallback authorization system.</span></span> <span data-ttu-id="c2412-265">Дополнительные сведения об этих ошибках см. в статье [Устранение ошибок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md). В этой надстройке резервная система открывает диалоговое окно, требующее входа пользователя, даже если он уже выполнил вход, и использует msal.js и неявный поток, чтобы получить маркер доступа к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="c2412-265">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is, and uses msal.js and the Implicit Flow to get an access token to Microsoft Graph.</span></span>

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a><span data-ttu-id="c2412-266">Обработка ошибок на стороне сервера</span><span class="sxs-lookup"><span data-stu-id="c2412-266">Handle server-side errors</span></span>

1. <span data-ttu-id="c2412-267">Добавьте указанную ниже функцию под функцией `handleClientSideErrors`.</span><span class="sxs-lookup"><span data-stu-id="c2412-267">Below the `handleClientSideErrors` function, add the following function.</span></span>

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. <span data-ttu-id="c2412-268">Замените `TODO 4` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-268">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="c2412-269">Вот что нужно знать об этом коде: классы ошибок в ASP.NET были созданы до появления MFA.</span><span class="sxs-lookup"><span data-stu-id="c2412-269">About this code, note that ASP.NET error classes were created before there was such a thing as MFA.</span></span> <span data-ttu-id="c2412-270">Побочным эффектом того, как логика на стороне сервера обрабатывает запросы второго фактора проверки подлинности, является то, что у ошибки на стороне сервера, отправляемой клиенту, есть свойство **Message**, но нет свойства **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="c2412-270">As a side-effect of how our server-side logic handles the requests for a second authentication factor, the server-side error sent to the client has a **Message** property but no **ExceptionMessage** property.</span></span> <span data-ttu-id="c2412-271">Однако у всех остальных ошибок будет свойство **ExceptionMessage**, поэтому клиентский код должен проанализировать ответ для обоих свойств. </span><span class="sxs-lookup"><span data-stu-id="c2412-271">But all other errors will have a **ExceptionMessage** property, so the client-side code has to parse the response for both.</span></span> <span data-ttu-id="c2412-272">Одна из переменных не будет определена.</span><span class="sxs-lookup"><span data-stu-id="c2412-272">Either one or the other variable will be undefined.</span></span>

    ```javascript
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. <span data-ttu-id="c2412-273">Замените `TODO 5` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-273">Replace `TODO 5` with the following.</span></span> <span data-ttu-id="c2412-274">Когда Microsoft Graph требует дополнительной проверки подлинности, он отправляет ошибку AADSTS50076.</span><span class="sxs-lookup"><span data-stu-id="c2412-274">When Microsoft Graph requires an additional form of authentication, it sends error AADSTS50076.</span></span> <span data-ttu-id="c2412-275">Она содержит сведения о дополнительном требовании в свойстве **Message.Claims**.</span><span class="sxs-lookup"><span data-stu-id="c2412-275">It includes information about the additional requirement in the **Message.Claims** property.</span></span> <span data-ttu-id="c2412-276">Чтобы обработать эту ошибку, код делает вторую попытку получить маркер начальной загрузки, но в этот раз он включает запрос дополнительного фактора в виде значения параметра `authChallenge`, который предписывает Azure AD предложить пользователю пройти все требуемые проверки подлинности. </span><span class="sxs-lookup"><span data-stu-id="c2412-276">To handle this, the code makes a second attempt to get the bootstrap token, but this time it includes the request for an additional factor as the value of the `authChallenge` option, which tells Azure AD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. <span data-ttu-id="c2412-277">Замените `TODO 6` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-277">Replace `TODO 6` with the following.</span></span>

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. <span data-ttu-id="c2412-278">Замените `TODO 7` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-278">Replace `TODO 7` with the following.</span></span> <span data-ttu-id="c2412-279">Обратите внимание, что иногда срок действия маркера начальной загрузки не истекает в момент его проверки в Office, но истекает ко времени его попадания в Azure AD для замены.</span><span class="sxs-lookup"><span data-stu-id="c2412-279">Note that on rare occasions the bootstrap token is unexpired when Office validates it, but expires by the time it is sent to Azure AD for exchange.</span></span> <span data-ttu-id="c2412-280">Служба Azure AD ответит ошибкой AADSTS500133.</span><span class="sxs-lookup"><span data-stu-id="c2412-280">Azure AD will respond with error AADSTS500133.</span></span> <span data-ttu-id="c2412-281">В этом случае код вызывает API единого входа (но не более одного раза).</span><span class="sxs-lookup"><span data-stu-id="c2412-281">When this happens, the code  recalls the SSO API (but no more than once).</span></span> <span data-ttu-id="c2412-282">На этот раз Office возвращает новый маркер начальной загрузки, срок действия которого не истек.  </span><span class="sxs-lookup"><span data-stu-id="c2412-282">This time Office returns a new unexpired bootstrap token.</span></span>

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="c2412-283">Замените `TODO 8` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-283">Replace `TODO 8` with the following.</span></span>

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. <span data-ttu-id="c2412-284">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="c2412-284">Save the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="c2412-285">Код на стороне сервера</span><span class="sxs-lookup"><span data-stu-id="c2412-285">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="c2412-286">Настройка ПО промежуточного слоя OWIN</span><span class="sxs-lookup"><span data-stu-id="c2412-286">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="c2412-287">Откройте файл Startup.cs в корневой папке проекта **Office-Add-in-ASPNET-SSO-WebAPI** и добавьте приведенный ниже метод в класс **Startup**.</span><span class="sxs-lookup"><span data-stu-id="c2412-287">Open the Startup.cs file in the root of the **Office-Add-in-ASPNET-SSO-WebAPI** project and add the following method to the **Startup** class.</span></span> <span data-ttu-id="c2412-288">Обратите внимание, что метод `ConfigureAuth` создается позже.</span><span class="sxs-lookup"><span data-stu-id="c2412-288">Note that you create the `ConfigureAuth` method in a later step.</span></span>

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. <span data-ttu-id="c2412-289">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="c2412-289">Save and close the file.</span></span>

1. <span data-ttu-id="c2412-290">Щелкните правой кнопкой мыши папку **App_Start** и выберите **Добавить > Класс**.</span><span class="sxs-lookup"><span data-stu-id="c2412-290">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="c2412-291">В диалоговом окне **Добавить новый элемент** введите имя файла **Startup.Auth.cs** и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="c2412-291">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="c2412-292">Сократите имя пространства имен в новом файле до `Office_Add_in_ASPNET_SSO_WebAPI`.</span><span class="sxs-lookup"><span data-stu-id="c2412-292">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="c2412-293">Убедитесь, что в начале файла есть все приведенные ниже операторы `using`.</span><span class="sxs-lookup"><span data-stu-id="c2412-293">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="c2412-p148">Добавьте ключевое слово `partial` в объявление класса `Startup`, если его там еще нет. Оно должно выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="c2412-p148">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="c2412-p149">Добавьте приведенный ниже метод в класс `Startup`. Этот метод указывает, как ПО промежуточного слоя OWIN будет проверять маркеры доступа, передаваемые ему из метода `getData` в файле Home.js на стороне клиента. Процесс вызывается при каждом вызове конечной точки веб-API, содержащей атрибут `[Authorize]`.</span><span class="sxs-lookup"><span data-stu-id="c2412-p149">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. <span data-ttu-id="c2412-299">Замените `TODO 1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-299">Replace the `TODO 1` with the following.</span></span> <span data-ttu-id="c2412-300">Что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="c2412-300">Note about this code:</span></span>

    * <span data-ttu-id="c2412-301">Код предписывает OWIN убедиться, что аудитория, указанная в маркере начальной загрузки из ведущего приложения Office, совпадает со значением, указанным в файле web.config.</span><span class="sxs-lookup"><span data-stu-id="c2412-301">The code instructs OWIN to ensure that the audience specified in the bootstrap token that comes from the Office host must match the value specified in the web.config.</span></span>
    * <span data-ttu-id="c2412-302">У учетных записей Майкрософт есть идентификатор GUID поставщика, отличный от GUID корпоративного клиента. Чтобы поддержать оба вида учетных записей, поставщик не проверяется.</span><span class="sxs-lookup"><span data-stu-id="c2412-302">Microsoft Accounts have an issuer GUID that is different from any organizational tenant GUID, so to support both kinds of accounts, we do not validate the issuer.</span></span>
    * <span data-ttu-id="c2412-303">Если задать для свойства `SaveSigninToken` значение `true`, OWIN сохранит необработанный маркер начальной загрузки из ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="c2412-303">Setting `SaveSigninToken` to `true` causes OWIN to save the raw bootstrap token from the Office host.</span></span> <span data-ttu-id="c2412-304">Он необходим надстройке, чтобы получить маркер доступа к Microsoft Graph в потоке "от имени".</span><span class="sxs-lookup"><span data-stu-id="c2412-304">The add-in needs it to obtain an access token to Microsoft Graph with the on-behalf-of flow.</span></span>
    * <span data-ttu-id="c2412-305">ПО промежуточного слоя OWIN не проверяет области.</span><span class="sxs-lookup"><span data-stu-id="c2412-305">Scopes are not validated by the OWIN middleware.</span></span> <span data-ttu-id="c2412-306">Области маркера начальной загрузки, которые должны включать `access_as_user`, проверяются в контроллере.</span><span class="sxs-lookup"><span data-stu-id="c2412-306">The scopes of the bootstrap token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. <span data-ttu-id="c2412-307">Замените `TODO 2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-307">Replace `TODO 2` with the following.</span></span> <span data-ttu-id="c2412-308">Что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="c2412-308">Note about this code:</span></span>

    * <span data-ttu-id="c2412-309">Метод `UseOAuthBearerAuthentication` вызывается вместо более распространенного метода `UseWindowsAzureActiveDirectoryBearerAuthentication`, так как последний несовместим с конечной точкой Azure AD версии 2.</span><span class="sxs-lookup"><span data-stu-id="c2412-309">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="c2412-310">ПО промежуточного слоя OWIN использует URL-адрес, передаваемый методу, чтобы получить ключ, необходимый для проверки подписи в маркере начальной загрузки, полученном из ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="c2412-310">The URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the bootstrap token received from the Office host.</span></span> <span data-ttu-id="c2412-311">Сегмент URL-адреса "Полномочия" предоставляется файлом web.config. Это либо строка "common", либо GUID для однотенантной надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2412-311">The Authority segment of the URL comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. <span data-ttu-id="c2412-312">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="c2412-312">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="c2412-313">Создание контроллера /api/values</span><span class="sxs-lookup"><span data-stu-id="c2412-313">Create the /api/values controller</span></span>

1. <span data-ttu-id="c2412-314">Откройте файл **Controllers\ValueController.cs**.</span><span class="sxs-lookup"><span data-stu-id="c2412-314">Open the file **Controllers\ValueController.cs**.</span></span> <span data-ttu-id="c2412-315">Этот контроллер используется в случае успешного получения маркера начальной загрузки системой единого входа.</span><span class="sxs-lookup"><span data-stu-id="c2412-315">This controller is used when the SSO system has successfully obtained a bootstrap token.</span></span> <span data-ttu-id="c2412-316">Он не используется в рамках резервной системы авторизации.</span><span class="sxs-lookup"><span data-stu-id="c2412-316">It is not used as part of the fallback authorization system.</span></span> <span data-ttu-id="c2412-317">В этой системе использован AzureADAuthController, созданный для вас.</span><span class="sxs-lookup"><span data-stu-id="c2412-317">That system used the AzureADAuthController, which has been created for you.</span></span>

1. <span data-ttu-id="c2412-318">Убедитесь, что в начале файла есть приведенные ниже инструкции с `using`.</span><span class="sxs-lookup"><span data-stu-id="c2412-318">Ensure that the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Microsoft.Identity.Client;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    ```

1. <span data-ttu-id="c2412-p156">Над строкой с объявлением `ValuesController` добавьте атрибут `[Authorize]`. Это гарантирует, что надстройка будет выполнять процесс авторизации, настроенный в последней процедуре, при каждом вызове метода контроллера. Вызывать методы контроллера можно только при наличии действительного маркера доступа к надстройке.</span><span class="sxs-lookup"><span data-stu-id="c2412-p156">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

1. <span data-ttu-id="c2412-322">Добавьте приведенный ниже метод в `ValuesController`.</span><span class="sxs-lookup"><span data-stu-id="c2412-322">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="c2412-323">Обратите внимание, что возвращаемое значение — `Task<HttpResponseMessage>`, а не `Task<IEnumerable<string>>`, которое чаще используется для метода `GET api/values`.</span><span class="sxs-lookup"><span data-stu-id="c2412-323">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="c2412-324">Это побочный эффект того, что логика авторизации OAuth находится в контроллере, а не в фильтре ASP.NET.</span><span class="sxs-lookup"><span data-stu-id="c2412-324">This is a side effect of that fact that the OAuth  authorization logic must be in the controller, instead of in an ASP.NET filter.</span></span> <span data-ttu-id="c2412-325">Некоторые условия возникновения ошибки в этой логике требуют отправки объекта HTTP-ответа в клиент надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2412-325">Some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span>

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //        token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get the access token for Microsoft Graph.

        // TODO 4: Use the token to call Microsoft Graph.
    }
    ```

1. <span data-ttu-id="c2412-326">Замените `TODO1` приведенным ниже кодом, чтобы убедиться, что в маркере указано разрешение `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="c2412-326">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span> <span data-ttu-id="c2412-327">Обратите внимание, что второй параметр метода `SendErrorToClient` — объект **Exception**.</span><span class="sxs-lookup"><span data-stu-id="c2412-327">Note that the second parameter of the `SendErrorToClient` method is an **Exception** object.</span></span> <span data-ttu-id="c2412-328">В этом случае код передает `null`, потому что включение объекта **Exception** блокирует включение свойства **Message** в создаваемый HTTP-ответ.</span><span class="sxs-lookup"><span data-stu-id="c2412-328">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>


    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. <span data-ttu-id="c2412-329">Замените `TODO 2` приведенным ниже кодом, чтобы собрать все сведения, необходимые для получения маркера для Microsoft Graph, используя поток "от имени".</span><span class="sxs-lookup"><span data-stu-id="c2412-329">Replace `TODO 2` with the following code to assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.</span></span> <span data-ttu-id="c2412-330">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="c2412-330">About this code, note:</span></span>

    * <span data-ttu-id="c2412-p160">Надстройка больше не выступает в роли ресурса (или аудитории), доступ к которому необходим ведущему приложению Office и пользователю. Теперь она сама является клиентом, которому необходим доступ к Microsoft Graph. `ConfidentialClientApplication` — это объект "контекста клиента" MSAL.</span><span class="sxs-lookup"><span data-stu-id="c2412-p160">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="c2412-334">Начиная с MSAL.NET 3.x.x, `bootstrapContext` — это сам маркер начальной загрузки. </span><span class="sxs-lookup"><span data-stu-id="c2412-334">Beginning with MSAL.NET 3.x.x, the `bootstrapContext` is just the bootstrap token itself.</span></span>
    * <span data-ttu-id="c2412-335">Полномочия предоставляются файлом web.config. Это либо строка "common", либо GUID для однотенантной надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2412-335">The Authority comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>
    * <span data-ttu-id="c2412-p161">Для работы библиотеки MSAL требуются области `openid` и `offline_access`, но если код их избыточно запрашивает, возникает ошибка. Кроме того, ошибка возникнет, если код запросит `profile` (фактически используется только при получении ведущим приложением Office токена для веб-приложения надстройки). Поэтому явным образом запрашивается только `Files.Read.All`.</span><span class="sxs-lookup"><span data-stu-id="c2412-p161">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them. It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application. So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri("https://localhost:44355")
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. <span data-ttu-id="c2412-p162">Замените `TODO 3` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="c2412-p162">Replace `TODO 3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="c2412-341">Для начала метод `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` проверит кэш MSAL, который находится в памяти, на наличие подходящего маркера доступа.</span><span class="sxs-lookup"><span data-stu-id="c2412-341">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token.</span></span> <span data-ttu-id="c2412-342">Только в случае его отсутствия запускается поток "от имени" с конечной точкой Azure AD версии 2.</span><span class="sxs-lookup"><span data-stu-id="c2412-342">Only if there isn't one, does it initiate the on-behalf-of flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="c2412-343">Любые исключения, отличные от типа `MsalServiceException`, не перехватываются преднамеренно, поэтому будут переданы клиенту в виде сообщений `500 Server Error`.</span><span class="sxs-lookup"><span data-stu-id="c2412-343">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

    ```csharp
    AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
    AuthenticationResult authResult = null;
    try
    {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
    }
    catch (MsalServiceException e)
    {
        // TODO 3a: Handle request for multi-factor authentication.

        // TODO 3b: Handle lack of consent and invalid scope (permission).

        // TODO 3c: Handle all other MsalServiceExceptions.
    }
    ```

1. <span data-ttu-id="c2412-344">Замените `TODO 3a` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-344">Replace `TODO 3a` with the following code.</span></span> <span data-ttu-id="c2412-345">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="c2412-345">About this code, note:</span></span>

    * <span data-ttu-id="c2412-346">Если ресурс Microsoft Graph требует многофакторной проверки подлинности, а пользователь еще не предоставил соответствующие данные, Azure AD вернет состояние "400 Bad Request" с ошибкой `AADSTS50076` и свойство **Claims**.</span><span class="sxs-lookup"><span data-stu-id="c2412-346">If multi-factor authentication is required by the Microsoft Graph resource and the user has not yet provided it, Azure AD will return "400 Bad Request" with error `AADSTS50076` and a **Claims** property.</span></span> <span data-ttu-id="c2412-347">MSAL выдает исключение **MsalUiRequiredException** (которое наследуется от **MsalServiceException**), используя эту информацию.</span><span class="sxs-lookup"><span data-stu-id="c2412-347">MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span>
    * <span data-ttu-id="c2412-348">Значение свойства **Claims** необходимо передать клиенту, который передаст его ведущему приложению Office. Последнее добавит его в запрос на получение нового маркера начальной загрузки.</span><span class="sxs-lookup"><span data-stu-id="c2412-348">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new bootstrap token.</span></span> <span data-ttu-id="c2412-349">Azure AD предложит пользователю пройти все необходимые проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="c2412-349">Azure AD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="c2412-p167">API, которые создают HTTP-ответы из исключений, не знают о свойстве **Claims**, поэтому не включают его в ответ. Нам нужно создать сообщение с ним вручную. Однако настраиваемое свойство **Message** блокирует создание свойства **ExceptionMessage**, поэтому единственный способ передать идентификатор ошибки `AADSTS50076` клиенту — добавить его в настраиваемое свойство **Message**. Код JavaScript в клиенте должен будет определить, какое свойство содержится в ответе (**Message** или **ExceptionMessage**).</span><span class="sxs-lookup"><span data-stu-id="c2412-p167">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="c2412-354">Сообщение создается в формате JSON, чтобы клиентский код JavaScript мог проанализировать его с помощью известных методов объекта JavaScript `JSON`.</span><span class="sxs-lookup"><span data-stu-id="c2412-354">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known JavaScript `JSON` object methods.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. <span data-ttu-id="c2412-355">Замените `TODO 3b` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-355">Replace `TODO 3b` with the following code.</span></span> <span data-ttu-id="c2412-356">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="c2412-356">About this code, note:</span></span>

    * <span data-ttu-id="c2412-357">Если вызов Azure AD содержал по крайней мере одно разрешение, которое не предоставил ни пользователь, ни администратор клиента (или оно было отозвано), Azure AD вернет состояние "400 Bad Request" с ошибкой `AADSTS65001`.</span><span class="sxs-lookup"><span data-stu-id="c2412-357">If the call to Azure AD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked), Azure AD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="c2412-358">MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию.</span><span class="sxs-lookup"><span data-stu-id="c2412-358">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    *  <span data-ttu-id="c2412-359">Если вызов Azure AD содержал по крайней мере одно нераспознанное разрешение, Azure AD вернет состояние "400 Bad Request" с ошибкой `AADSTS70011`.</span><span class="sxs-lookup"><span data-stu-id="c2412-359">If the call to Azure AD contained at least one scope that Azure AD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="c2412-360">MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию.</span><span class="sxs-lookup"><span data-stu-id="c2412-360">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    *  <span data-ttu-id="c2412-361">Полное описание включается, так как ошибка 70011 возвращается и в других случаях, и ее следует обрабатывать в этой надстройке, только когда она означает запрос недопустимого разрешения.</span><span class="sxs-lookup"><span data-stu-id="c2412-361">The entire description is included because 70011 is returned in other conditions and it should only be handled in this add-in when it means that there is an invalid scope.</span></span>
    *  <span data-ttu-id="c2412-p171">Объект **MsalUiRequiredException** передается методу `SendErrorToClient`. Это гарантирует, что свойство **ExceptionMessage**, содержащее информацию об ошибке, будет включено в HTTP-отклик.</span><span class="sxs-lookup"><span data-stu-id="c2412-p171">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. <span data-ttu-id="c2412-364">Замените `TODO 3c` приведенным ниже кодом, чтобы обработать все остальные исключения **MsalServiceException**.</span><span class="sxs-lookup"><span data-stu-id="c2412-364">Replace `TODO 3c` with the following code to handle all other **MsalServiceException**s.</span></span> <span data-ttu-id="c2412-365">Как отмечалось выше,</span><span class="sxs-lookup"><span data-stu-id="c2412-365">As noted earlier,</span></span>

    ```csharp
    else
    {
        throw e;
    }
    ```

1. <span data-ttu-id="c2412-366">замените `TODO 4` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="c2412-366">Replace `TODO 4` with the following code.</span></span> <span data-ttu-id="c2412-367">Метод `GraphApiHelper.GetOneDriveFileNames`, созданный для вас, выполняет запрос данных в Microsoft Graph и включает маркер доступа.</span><span class="sxs-lookup"><span data-stu-id="c2412-367">The `GraphApiHelper.GetOneDriveFileNames` method, which has been created for you, makes the request for data to Microsoft Graph and includes the access token.</span></span>

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. <span data-ttu-id="c2412-368">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="c2412-368">Save and close the file.</span></span>

## <a name="run-the-solution"></a><span data-ttu-id="c2412-369">Запуск решения</span><span class="sxs-lookup"><span data-stu-id="c2412-369">Run the solution</span></span>

1. <span data-ttu-id="c2412-370">Откройте файл решения в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="c2412-370">Open the Visual Studio solution file.</span></span>
1. <span data-ttu-id="c2412-371">В меню **Построение** выберите команду **Очистить решение**.</span><span class="sxs-lookup"><span data-stu-id="c2412-371">On the **Build** menu, select **Clean Solution**.</span></span> <span data-ttu-id="c2412-372">После выполнения команды снова откройте меню **Построение** и выберите команду **Построить решение**.</span><span class="sxs-lookup"><span data-stu-id="c2412-372">When it finishes, open the **Build** menu again and select **Build Solution**.</span></span>
1. <span data-ttu-id="c2412-373">В **обозревателе решений** выберите узел проекта **Office-Add-in-ASPNET-SSO** (не верхний узел решения и не узел проекта, имя которого заканчивается на "WebAPI").</span><span class="sxs-lookup"><span data-stu-id="c2412-373">In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO** project node (not the top solution node and not the project whose name ends in "WebAPI").</span></span>
1. <span data-ttu-id="c2412-374">В области **Свойства** откройте раскрывающийся список **Начальный документ** и выберите один из трех вариантов (Excel, Word или PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="c2412-374">In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).</span></span>

    ![Выбор ведущего приложения Office: Excel, PowerPoint или Word](../images/SelectHost.JPG)

1. <span data-ttu-id="c2412-376">Нажмите клавишу F5.</span><span class="sxs-lookup"><span data-stu-id="c2412-376">Press F5.</span></span>
1. <span data-ttu-id="c2412-377">В приложении Office на вкладке ленты **Главная** в группе **Единый вход ASP.NET** выберите команду **Показать надстройку**, чтобы открыть надстройку области задач.</span><span class="sxs-lookup"><span data-stu-id="c2412-377">In the Office application, on the **Home** ribbon, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.</span></span>
1. <span data-ttu-id="c2412-378">Нажмите кнопку **Получить имена файлов OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="c2412-378">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="c2412-379">Если вы выполнили вход в Office с помощью рабочей или учебной учетной записи (Office 365) либо учетной записи Майкрософт и единый вход работает надлежащим образом, первые 10 имен файлов и папок из OneDrive для бизнеса отобразятся в области задач.</span><span class="sxs-lookup"><span data-stu-id="c2412-379">If you are logged into Office with either a Work or School (Office 365) account or Microsoft Account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane.</span></span> <span data-ttu-id="c2412-380">Если вы не выполнили вход или используете сценарий, не поддерживающий единый вход, или единый вход не работает по какой-то причине, появится запрос на вход.</span><span class="sxs-lookup"><span data-stu-id="c2412-380">If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in.</span></span> <span data-ttu-id="c2412-381">После входа в систему отобразятся имена файлов и папок.</span><span class="sxs-lookup"><span data-stu-id="c2412-381">After you log in, the file and folder names appear.</span></span>
