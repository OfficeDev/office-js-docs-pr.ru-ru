---
title: Создание надстройки Office на платформе Node.js с использованием единого входа
description: Узнайте, как создать надстройку на основе Node.js, использующую единый вход Office
ms.date: 11/20/2019
localization_priority: Priority
ms.openlocfilehash: 362ca4a534800a683284b049e6e53776b1aa7f38
ms.sourcegitcommit: 013886c1b08ef2b378cf80bb88bc73ec56c3e869
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/22/2019
ms.locfileid: "39191741"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="9c4b3-103">Создание надстройки Office на платформе Node.js с использованием единого входа (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="9c4b3-103">Create a Node.js Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="9c4b3-p101">Ваша веб-надстройка Office может использовать процедуру входа в Office для авторизации пользователей в надстройке и Microsoft Graph. При этом им не потребуется входить повторно. Общие сведения см. в статье [Включение единого входа в надстройке Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="9c4b3-106">Из этой статьи вы узнаете, как включить единый вход в надстройке, созданной с помощью Node.js и Express.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span>

> [!NOTE]
> <span data-ttu-id="9c4b3-107">Аналогичная статья, посвященная надстройке на основе ASP.NET, — [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-107">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="9c4b3-108">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="9c4b3-108">Prerequisites</span></span>

* <span data-ttu-id="9c4b3-109">[Node и npm](https://nodejs.org/) версии 10.15.0 или более поздней.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-109">[Node and npm](https://nodejs.org/), version 10.15.0 or later.</span></span>

* <span data-ttu-id="9c4b3-110">[Git Bash](https://git-scm.com/downloads) (или другой клиент git).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-110">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="9c4b3-111">TypeScript версии 3.6.2 или более поздней.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-111">TypeScript, version 3.6.2 or later</span></span>

* <span data-ttu-id="9c4b3-112">Учетная запись Office 365 (версия Office с подпиской), которую получают участники [программы для разработчиков Office 365](https://aka.ms/devprogramsignup), предоставляется вместе с бесплатной годичной подпиской на Office 365.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-112">An Office 365 account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365.</span></span> <span data-ttu-id="9c4b3-113">Чтобы получить эту версию, следует использовать последнюю ежемесячную версию и сборку из канала программы предварительной оценки, при этом необходимо быть участником программы предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-113">You should use the latest monthly version and build from the Insiders channel but you need to be an Office Insider to get this version.</span></span> <span data-ttu-id="9c4b3-114">Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-114">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="9c4b3-115">Обратите внимание на то, что когда сборка будет готова для выпуска на канале Semi-annual channel, поддержка функций предварительного просмотра, включая единый вход, отключается для этой сборки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-115">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

* <span data-ttu-id="9c4b3-116">Редактор кода.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-116">A source code editor.</span></span> <span data-ttu-id="9c4b3-117">Рекомендуется использовать Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-117">We recommend Visual Studio Code.</span></span>

* <span data-ttu-id="9c4b3-118">Несколько файлов и папок, сохраненных в OneDrive для бизнеса в составе подписки на Office 365.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-118">At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.</span></span>

* <span data-ttu-id="9c4b3-119">Клиент Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-119">A Microsoft Azure Tenant to register your application.</span></span> <span data-ttu-id="9c4b3-120">Эта надстройка требует наличия Azure Active Directory (AD).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-120">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="9c4b3-121">В Azure AD доступны службы идентификации, которые приложения используют для проверки подлинности и авторизации.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-121">Azure Active Directory provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="9c4b3-122">Пробную подписку можно получить на сайте [Microsoft Azure](https://account.windowsazure.com/SignUp).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-122">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="9c4b3-123">Настройка начального проекта</span><span class="sxs-lookup"><span data-stu-id="9c4b3-123">Set up the starter project</span></span>

1. <span data-ttu-id="9c4b3-124">Клонируйте или скачайте репозиторий [Office-Add-in-NodeJS-SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-124">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span>

    > [!NOTE]
    > <span data-ttu-id="9c4b3-125">Существует три версии примера.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-125">There are three versions of the sample:</span></span>  
    > * <span data-ttu-id="9c4b3-p105">В папке **Before** находится начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы. В последующих разделах этой статьи рассматривается доработка проекта.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-p105">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
    > * <span data-ttu-id="9c4b3-129">Версия примера в папке **Complete** идентична надстройке, которую вы бы создали, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-129">The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections Code the client side and Code the server side.</span></span> <span data-ttu-id="9c4b3-130">Чтобы использовать готовую версию, просто выполните действия, описанные в этой статье, но замените папку Before на папку Completed и пропустите разделы **Код на стороне клиента** и **Код на стороне сервера**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-130">The Completed version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="9c4b3-131">Версия **SSOAutoSetup** — это готовый пример, который автоматизирует большинство шагов регистрации надстройки в Azure AD и ее настройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-131">The **SSOAutoSetup** version is a completed sample that automates most of the steps to register the add-in with Azure AD and configure it.</span></span> <span data-ttu-id="9c4b3-132">Используйте эту версию, если нужно быстро получить рабочую надстройку с единым входом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-132">Use this version if you want to see a working add-in with SSO quickly.</span></span> <span data-ttu-id="9c4b3-133">Просто следуйте инструкциям файла сведений в папке.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-133">Just follow the steps in the Readme of the folder.</span></span> <span data-ttu-id="9c4b3-134">На определенном этапе рекомендуется выполнить шаги ручной регистрации и настройки из этой статьи, чтобы лучше понять связь между Azure AD и надстройкой.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-134">We recommend that at some point you go through the manual registration and setup steps in this article to better understand the relationship between Azure AD and an add-in.</span></span> 


1. <span data-ttu-id="9c4b3-135">Откройте командную строку в папке **Before**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-135">Open a command prompt in the **Before** folder.</span></span>

1. <span data-ttu-id="9c4b3-136">Введите в консоли команду `npm install`, чтобы установить все зависимости, указанные в файле package.json.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-136">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

1. <span data-ttu-id="9c4b3-137">Выполните команду `npm run install-dev-certs`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-137">Run the command  `npm run install-dev-certs`.</span></span> <span data-ttu-id="9c4b3-138">При запросе нажмите **Да**, чтобы установить сертификат.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-138">Select **Yes** to the prompt to disable the designer.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="9c4b3-139">Регистрация надстройки в конечной точке Azure AD версии 2.0</span><span class="sxs-lookup"><span data-stu-id="9c4b3-139">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="9c4b3-140">Перейдите на страницу [регистрации приложений портала Azure](https://go.microsoft.com/fwlink/?linkid=2083908), чтобы зарегистрировать свое приложение.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-140">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="9c4b3-141">Войдите в клиент Office 365, используя учетные данные ***администратора***.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-141">Sign in with the ***admin*** credentials to your Office 365 tenancy.</span></span> <span data-ttu-id="9c4b3-142">Пример: MyName@contoso.onmicrosoft.com.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-142">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="9c4b3-143">Выберите **Новая регистрация**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-143">Select **New registration**.</span></span> <span data-ttu-id="9c4b3-144">На странице**Зарегистрировать приложение** задайте необходимые значения следующим образом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-144">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="9c4b3-145">Введите **имя** `Office-Add-in-NodeJS-SSO`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-145">Set **Name** to `Office-Add-in-NodeJS-SSO`.</span></span>
    * <span data-ttu-id="9c4b3-146">Для параметра **Поддерживаемые типы учетных записей** укажите вариант **Учетные записи в любом каталоге организации и личные учетные записи Майкрософт (например, Skype, Xbox, Outlook.com)**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-146">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span>
    * <span data-ttu-id="9c4b3-147">Присвойте параметру **URI перенаправления** значение ` https://localhost:44355/dialog.html`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-147">Set **Redirect URI** to` https://localhost:44355/dialog.html`.</span></span>
    * <span data-ttu-id="9c4b3-148">Нажмите кнопку **Зарегистрировать**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-148">Choose **Register**.</span></span>

1. <span data-ttu-id="9c4b3-149">На странице **Office-Add-in-NodeJS-SSO** скопируйте и сохраните значения параметров **Идентификатор приложения (клиент)** и **Идентификатор каталога (клиент)**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-149">On the **$ADD-IN-NAME$** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="9c4b3-150">Они понадобятся вам позже.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-150">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9c4b3-151">Этот идентификатор представляет собой значение аудитории, используемое, когда другие приложения, например ведущее приложение Office (PowerPoint, Word, Excel и т. д.), пытаются получить авторизованный доступ к вашему приложению.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-151">This ID is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="9c4b3-152">Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-152">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="9c4b3-153">Выберите **Проверка подлинности** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-153">Select **Authentication** under **Manage**.</span></span> <span data-ttu-id="9c4b3-154">В разделе **Неявное представление** установите флажки **Маркер доступа** и **Токен идентификатора**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-154">In the **Implict grant** section, enable the checkboxes for both **Access token** and **ID token**.</span></span> <span data-ttu-id="9c4b3-155">В примере используется резервная система авторизации, вызываемая при недоступности единого входа.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-155">The sample has a fallback authorization system that is invoked when SSO is not available.</span></span> <span data-ttu-id="9c4b3-156">В этой системе используется неявный поток.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-156">This system uses the Implicit Flow.</span></span>

1. <span data-ttu-id="9c4b3-157">Щелкните **Сохранить** в верхней части формы.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-157">Select **Save** at the top of the form.</span></span>

1. <span data-ttu-id="9c4b3-158">Выберите **Сертификаты и секреты** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-158">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="9c4b3-159">Нажмите кнопку **Новый секрет клиента**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-159">Select the **New client secret** button.</span></span> <span data-ttu-id="9c4b3-160">Введите значение параметра **Описание**, выберите соответствующий вариант для параметра **Истекает срок действия** и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-160">Enter a value for **Description** then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="9c4b3-161">*Сразу скопируйте значение секрета клиента и сохраните его с идентификатором приложения* перед продолжением, так как он понадобится вам позже.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-161">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="9c4b3-162">Выберите пункт **Предоставление API** в разделе **Управление**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-162">Select **Expose an API** under **Manage**.</span></span> <span data-ttu-id="9c4b3-163">Щелкните ссылку **Задать**, чтобы создать URI идентификатора приложения в формате "api://$ИД приложения GUID$", где $App ID GUID$ — **идентификатор приложения (клиента)**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-163">Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span> <span data-ttu-id="9c4b3-164">Вставьте `localhost:44355/` (обратите внимание на косую черту "/" в конце) между двойной косой чертой и GUID.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-164">Insert the `localhost:44355/` (with a forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="9c4b3-165">Весь идентификатор должен отображаться в формате `api://localhost:44355/$App ID GUID$`, например: `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-165">The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span> 

1. <span data-ttu-id="9c4b3-166">Нажмите кнопку **Добавить область**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-166">Select the **Add a scope** button.</span></span> <span data-ttu-id="9c4b3-167">В открывшейся панели введите `access_as_user` в качестве параметра **Имя области**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-167">In the panel that opens, enter `access_as_user` as the **Scope name**.</span></span>

1. <span data-ttu-id="9c4b3-168">Для параметра **Кто может давать согласие?** установите вариант **Администраторы и пользователи**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-168">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="9c4b3-169">Заполните поля для настройки запросов согласия администраторов и пользователей значениями, соответствующими области `access_as_user`, позволяющей ведущему приложению Office использовать веб-интерфейсы API надстройки с такими же правами, как у текущего пользователя.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-169">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office host application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="9c4b3-170">Возможные варианты:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-170">Suggestions:</span></span>

    - <span data-ttu-id="9c4b3-171">**Отображаемое имя согласия администратора**. Office может действовать в качестве пользователя.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-171">**Admin consent title:** Office can act as the user.</span></span>
    - <span data-ttu-id="9c4b3-172">**Описание согласия администратора**. Позволяет Office вызывать веб-API надстройки с такими же правами, как у текущего пользователя.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-172">**Admin consent description:** Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    - <span data-ttu-id="9c4b3-173">**Отображаемое имя согласия пользователя**. Office может действовать от вашего имени.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-173">**User consent title:** Office can act as you.</span></span>
    - <span data-ttu-id="9c4b3-174">**Описание согласия администратора**. Позволяет Office вызывать веб-API надстройки с такими же правами, как у вас.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-174">**Admin consent description:** Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="9c4b3-175">Убедитесь, что параметру **Состояние** присвоено значение **Включено**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-175">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="9c4b3-176">Нажмите кнопку **Добавить область**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-176">Select **Add scope**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9c4b3-177">Доменная часть имени **области**, отображаемая непосредственно под текстовым полем, должна автоматически соответствовать URI идентификатора приложения, заданного ранее, с добавлением `/access_as_user` в конце, например: `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-177">The domain part of the Scope name displayed just below the text field should automatically match the Application ID URI set in the previous step, with  appended to the end; for example, .</span></span>

1. <span data-ttu-id="9c4b3-178">В разделе **Авторизованные клиентские приложения** укажите приложения, которые необходимо авторизовать для веб-приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-178">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="9c4b3-179">Необходимо обеспечить предварительную авторизацию для всех указанных ниже идентификаторов.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-179">Each of the following IDs needs to be pre-authorized.</span></span>

    - <span data-ttu-id="9c4b3-180">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-180">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="9c4b3-181">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-181">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="9c4b3-182">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-182">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="9c4b3-183">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-183">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office on the web)</span></span>

    <span data-ttu-id="9c4b3-184">Для каждого идентификатора сделайте следующее:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-184">For each ID, take these steps:</span></span>

    <span data-ttu-id="9c4b3-185">а)</span><span class="sxs-lookup"><span data-stu-id="9c4b3-185">a.</span></span> <span data-ttu-id="9c4b3-186">Нажмите кнопку **Добавить клиентское приложение**, в открывшейся панели присвойте параметру "Идентификатор клиента" соответствующий код GUID и установите флажок `api://localhost:44355/$App ID GUID$/access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-186">Select **Add a client application** button then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="9c4b3-187">б)</span><span class="sxs-lookup"><span data-stu-id="9c4b3-187">b.</span></span> <span data-ttu-id="9c4b3-188">Нажмите кнопку **Добавить приложение**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-188">Select **Add application**.</span></span>

1. <span data-ttu-id="9c4b3-189">Выберите пункт **Разрешения API** в разделе **Управление** и нажмите кнопку **Добавить разрешение**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-189">Select **API permissions** under **Manage** and select **Add a permission**.</span></span> <span data-ttu-id="9c4b3-190">В открывшейся панели выберите **Microsoft Graph** и щелкните **Делегированные разрешения**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-190">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="9c4b3-191">Используйте поле поиска **Выбрать разрешения**, чтобы найти нужные разрешения для надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-191">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="9c4b3-192">Выберите следующие параметры.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-192">Select from the following:</span></span> <span data-ttu-id="9c4b3-193">Для самой надстройки требуется только первое разрешение, но разрешение `profile` необходимо, чтобы ведущее приложение Office получило маркер для веб-приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-193">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>

    * <span data-ttu-id="9c4b3-194">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="9c4b3-194">Files.Read.All</span></span>
    * <span data-ttu-id="9c4b3-195">profile</span><span class="sxs-lookup"><span data-stu-id="9c4b3-195">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="9c4b3-196">Разрешение `User.Read` может быть уже указано по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-196">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="9c4b3-197">Незачем запрашивать ненужные разрешения, поэтому рекомендуем снять флажок рядом с разрешением, которое не требуется вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-197">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="9c4b3-198">Установите флажок для каждого отображаемого разрешения.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-198">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="9c4b3-199">Выбрав нужные для надстройки разрешения, нажмите кнопку **Добавить разрешения** в нижней части панели.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-199">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="9c4b3-200">На этой же странице нажмите кнопку **Предоставить согласие администратора для [имя клиента]** и выберите **Да** в появившемся запросе подтверждения.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-200">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears.</span></span>

## <a name="configure-the-add-in"></a><span data-ttu-id="9c4b3-201">Настройка надстройки</span><span class="sxs-lookup"><span data-stu-id="9c4b3-201">Configure the add-in</span></span>

1. <span data-ttu-id="9c4b3-202">Откройте папку `\Begin` в скопированном проекте в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-202">Open the `\Begin` folder in the cloned project in your code editor.</span></span>

1. <span data-ttu-id="9c4b3-203">Откройте файл `.ENV` и используйте значения, скопированные ранее.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-203">Open the `.ENV` file and use the values that you copied earlier.</span></span> <span data-ttu-id="9c4b3-204">Присвойте параметру **CLIENT_ID** значение вашего **идентификатора приложения (клиента)**, а параметру **CLIENT_SECRET** — значение секрета вашего клиента.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-204">Set the **CLIENT_ID** to your **Application (client) ID**, and set the **CLIENT_SECRET** to your client secret.</span></span> <span data-ttu-id="9c4b3-205">Значения **не** должны быть заключены в кавычки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-205">The values should **not** be in quotation marks.</span></span> <span data-ttu-id="9c4b3-206">По завершении файл должен выглядеть следующим образом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-206">When you are done, the file should be similar to the following:</span></span> 

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. <span data-ttu-id="9c4b3-207">Откройте файл `\public\javascripts\fallbackAuthDialog.js`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-207">Open the `\public\javascripts\fallbackAuthDialog.js` file.</span></span> <span data-ttu-id="9c4b3-208">В объявлении `msalConfig` замените заполнитель $application_GUID here$ на идентификатор приложения, скопированный во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-208">In the `msalConfig` declaration, replace the placeholder $application_GUID here$ with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="9c4b3-209">Значение не должно быть заключено в кавычки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-209">The value you specify must be enclosed in quotation marks (").</span></span>

1. <span data-ttu-id="9c4b3-210">Откройте файл манифеста надстройки manifest\manifest_local.xml и прокрутите его до конца.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-210">Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file.</span></span> <span data-ttu-id="9c4b3-211">Над закрывающим тегом `</VersionOverrides>` вы найдете следующую часть кода:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-211">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="9c4b3-212">Замените заполнитель "$application_GUID here$" *в обоих местах* разметки идентификатором приложения, скопированным при регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-212">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="9c4b3-213">Символы "$" не входят в состав идентификатора, их не нужно вставлять.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-213">The "" are not part of the ID, so do not include them.</span></span> <span data-ttu-id="9c4b3-214">Это тот же идентификатор, который использовался для ClientID и Audience в файле web.config.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-214">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9c4b3-215">Значение **Resource** — это **URI идентификатора приложения**, указанный при регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-215">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span> <span data-ttu-id="9c4b3-216">Раздел **Scopes** используется для создания диалогового окна согласия, только если надстройка продается в AppSource.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-216">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="9c4b3-217">Код на стороне клиента</span><span class="sxs-lookup"><span data-stu-id="9c4b3-217">Code the client-side</span></span>

### <a name="create-the-sso-logic"></a><span data-ttu-id="9c4b3-218">Создание логики единого входа</span><span class="sxs-lookup"><span data-stu-id="9c4b3-218">Create the SSO logic</span></span>

1. <span data-ttu-id="9c4b3-219">Откройте файл `public\javascripts\ssoAuthES6.js` в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-219">In your code editor, open the src\server.ts file.</span></span> <span data-ttu-id="9c4b3-220">В нем уже есть код, обеспечивающий поддержку обещаний (даже в Internet Explorer 11), и вызов `Office.onReady` для назначения обработчика единственной кнопки надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-220">It already has code that ensures that Promises are supported, even in Internet Explorer 11, and an `Office.onReady` call to assign a handler to the add-in's only button.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9c4b3-221">Как следует из названия, ssoAuthES6.js использует синтаксис JavaScript ES6, так как применение `async` и `await` хорошо демонстрирует простоту API единого входа.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-221">As the name suggests, the ssoAuthES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API.</span></span> <span data-ttu-id="9c4b3-222">После запуска сервера localhost этот файл будет преобразован в синтаксис ES5, чтобы пример запускался в Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-222">When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will run in Internet Explorer 11.</span></span> 

1. <span data-ttu-id="9c4b3-223">Добавьте следующий код под методом Office.onReady:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-223">Add the following code below the Office.onReady method:</span></span>

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exhange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         OfficeRuntime.auth.getAccessToken call.

        }
    }
    ```

1. <span data-ttu-id="9c4b3-224">Замените `TODO 1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-224">Replace `TODO 1` with the following code.</span></span> <span data-ttu-id="9c4b3-225">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-225">About this code, note:</span></span>

    - <span data-ttu-id="9c4b3-226">`OfficeRuntime.auth.getAccessToken` предписывает Office получить маркер начальной загрузки из Azure AD.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-226">`OfficeRuntime.auth.getAccessToken` instructs Office to get a bootstrap token from Azure AD.</span></span> <span data-ttu-id="9c4b3-227">Маркер начальной загрузки аналогичен маркеру идентификатора, но имеет свойство `scp` (scope) со значением `access-as-user`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-227">A bootstrap token is similar to an ID token, but it has a `scp` (scope) property with the value `access-as-user`.</span></span> <span data-ttu-id="9c4b3-228">Такой тип маркера веб-приложение может заменить на маркер доступа к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-228">This kind of token can be exchanged by a web application for an access token to Microsoft Graph.</span></span>
    - <span data-ttu-id="9c4b3-229">Если параметру `allowSignInPrompt` присвоено значение true, значит при отсутствии входа пользователя Office откроет всплывающее окно входа.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-229">Setting the `allowSignInPrompt`option to true means that if no user is currently signed into Office, then Office will open a popup sign-in prompt.</span></span>
    - <span data-ttu-id="9c4b3-230">Присвоение параметру `forMSGraphAccess` значения true сигнализирует Office, что надстройка планирует использовать маркер начальной загрузки для получения маркера доступа к Micrsoft Graph вместо его использования в качестве маркера идентификатора.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-230">Setting the `forMSGraphAccess` option to true signals to Office that the add-in intends to use the bootstrap token to get an access token to Micrsoft Graph, instead of just using it as an ID token.</span></span> <span data-ttu-id="9c4b3-231">Если администратор клиента не предоставил согласие на доступ надстройки к Microsoft Graph, `OfficeRuntime.auth.getAccessToken` возвращает ошибку **13012**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-231">If the tenant administrator has not granted consent to the add-in's access to Microsoft Graph, then `OfficeRuntime.auth.getAccessToken` returns error **13012**.</span></span> <span data-ttu-id="9c4b3-232">Надстройка может отреагировать переходом на альтернативную систему проверки подлинности. Это необходимо, так как Office может запрашивать согласие только на доступ к профилю пользователя Azure AD, а не к областям Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-232">The add-in can respond by falling back to an alternative system of authorization, which is necessary because Office can prompt only for consent to the user's Azure AD profile, not to any Microsoft Graph scopes.</span></span> <span data-ttu-id="9c4b3-233">Резервная система проверки подлинности требует повторного входа пользователя в систему, и у пользователя *может* быть запрошено согласие на доступ к областям Micrsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-233">The fallback authorization system requires the user to sign in again and the user *can* be prompted to consent to Micrsoft Graph scopes.</span></span> <span data-ttu-id="9c4b3-234">Таким образом, параметр `forMSGraphAccess` обеспечивает, что надстройка не будет выполнять замену маркера, которая завершится ошибкой из-за отсутствия согласия.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-234">So, the `forMSGraphAccess` option ensures that the add-in won't make a token exchange that will fail due to lack of consent.</span></span> <span data-ttu-id="9c4b3-235">(Так как вы предоставили согласие администратора на предыдущем шаге, этот сценарий не возникнет для этой надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-235">(Since you granted administrator consent in an earlier step, this scenario won't happen for this add-in.</span></span> <span data-ttu-id="9c4b3-236">Но этот параметр добавлен в любом случае, чтобы продемонстрировать рекомендацию.)</span><span class="sxs-lookup"><span data-stu-id="9c4b3-236">But the option is included here anyway to illustrate a best practice.)</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true }); 
    ```

1. <span data-ttu-id="9c4b3-237">Замените `TODO 2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-237">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="9c4b3-238">Вы создадите метод `getGraphToken` на одном из следующих шагов.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-238">You'll create the `getGraphToken` method in a later step.</span></span>

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. <span data-ttu-id="9c4b3-239">Замените `TODO 3` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-239">Replace `TODO 3` with the following.</span></span> <span data-ttu-id="9c4b3-240">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-240">About this code, note:</span></span> 

    - <span data-ttu-id="9c4b3-241">Если клиент Office 365 настроен на обязательное применение многофакторной проверки подлинности, в параметр `exchangeResponse` будет включено свойство `claims` со сведениями о дополнительных обязательных факторах.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-241">If the Office 365 tenant has been configured to require multifactor authentication, then the `exchangeResponse` will include a `claims` property with information about the additional required factors.</span></span> <span data-ttu-id="9c4b3-242">В этом случае следует снова вызвать `OfficeRuntime.auth.getAccessToken` с присвоением параметру `authChallenge` значения свойства утверждений.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-242">In that case, `OfficeRuntime.auth.getAccessToken` should be called again with the `authChallenge` option set to the value of the claims property.</span></span> <span data-ttu-id="9c4b3-243">В результате AAD предложит пользователю пройти все необходимые проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-243">This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. <span data-ttu-id="9c4b3-244">Замените `TODO 4` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-244">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="9c4b3-245">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-245">About this code, note:</span></span> 

    - <span data-ttu-id="9c4b3-246">Вы создадите метод `handleAADErrors` на одном из следующих шагов.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-246">You'll create the `handleAADErrors` method in a later step.</span></span> <span data-ttu-id="9c4b3-247">Ошибки Azure AD возвращаются клиенту в виде откликов HTTP с кодом 200.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-247">Azure AD errors are returned to the client as HTTP code 200 Responses.</span></span> <span data-ttu-id="9c4b3-248">Они не вызывают ошибки, поэтому не запускается блок `catch` метода `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-248">They do not throw errors, so they do not trigger the `catch` block of the `getGraphData` method.</span></span>
    - <span data-ttu-id="9c4b3-249">Вы создадите метод `makeGraphApiCall` на одном из следующих шагов.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-249">You'll create the `makeGraphApiCall` method in a later step.</span></span> <span data-ttu-id="9c4b3-250">Он выполняет вызов AJAX к конечной точке MS Graph.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-250">It makes an AJAX call to the MS Graph endpoint.</span></span> <span data-ttu-id="9c4b3-251">Ошибки перехватываются обратным вызовом `.fail` этого вызова, а не блоком `catch` метода `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-251">Errors are caught in the `.fail` callback of that call, not in the `catch` block of the `getGraphData` method.</span></span>

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. <span data-ttu-id="9c4b3-252">Замените `TODO 5` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-252">Replace `TODO 5` with the following.</span></span>

    - <span data-ttu-id="9c4b3-253">Ошибки вызова `getAccessToken` будут иметь свойство `code` с номером ошибки (обычно в диапазоне 13xxx).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-253">Errors from the call of `getAccessToken` will have a `code` property with an error number, typically in the 13xxx range.</span></span> <span data-ttu-id="9c4b3-254">Вы создадите метод `handleClientSideErrors` на одном из следующих шагов.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-254">You'll create the `handleClientSideErrors` method in a later step.</span></span>
    - <span data-ttu-id="9c4b3-255">Метод `showMessage` отображает текст на панели задач.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-255">The `showMessage` method displays text on the task pane.</span></span>

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. <span data-ttu-id="9c4b3-256">Под методом `getGraphData` добавьте следующую функцию.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-256">Below the `getGraphData` method, add the following.</span></span> <span data-ttu-id="9c4b3-257">Обратите внимание, что `/auth` — это серверный экспресс-маршрут, заменяющий маркер начальной загрузки в Azure AD на маркер доступа к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-257">Note that `/auth` is a server-side Express route that exhanges the bootstrap token with Azure AD for an access token to Microsoft Graph.</span></span>

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. <span data-ttu-id="9c4b3-258">Под методом `getGraphToken` добавьте следующую функцию.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-258">Below the `getGraphToken` method, add the following.</span></span> <span data-ttu-id="9c4b3-259">Обратите внимание, что `error.code` — это число (обычно в диапазоне 13xxx).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-259">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```
1. <span data-ttu-id="9c4b3-260">Замените `TODO 6` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-260">Replace `TODO 6` with the following code.</span></span> <span data-ttu-id="9c4b3-261">Дополнительные сведения об этих ошибках см. в статье [Устранение ошибок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-261">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span> 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the Web.
        showMessage("Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The OfficeRuntime.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. <span data-ttu-id="9c4b3-262">Замените `TODO 7` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-262">Replace `TODO 7` with the following code.</span></span> <span data-ttu-id="9c4b3-263">Дополнительные сведения об этих ошибках см. в статье [Устранение ошибок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md). Функция `dialogFallback` вызывает альтернативную систему проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-263">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). The function `dialogFallback` invokes the alternative system of authorization.</span></span> <span data-ttu-id="9c4b3-264">В этой надстройке резервная система открывает диалоговое окно, требующее входа пользователя, даже если он уже выполнил вход, и использует msal.js и неявный поток, чтобы получить маркер доступа к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-264">In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is, and uses msal.js and the Implicit Flow to get an access token to Microsoft Graph.</span></span>

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. <span data-ttu-id="9c4b3-265">Добавьте указанную ниже функцию под функцией `handleClientSideErrors`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-265">Below the `handleClientSideErrors` function, add the following function.</span></span> 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. <span data-ttu-id="9c4b3-266">Иногда срок действия маркера начальной загрузки, кэшированного в Office, не истекает в момент его проверки в Office, но истекает ко времени его попадания в Azure AD для замены.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-266">On rare occasions the bootstrap token that Office has cached is unexpired when Office validates it, but expires by the time it reaches Azure AD for exchange.</span></span> <span data-ttu-id="9c4b3-267">Служба Azure AD ответит ошибкой **AADSTS500133**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-267">Azure AD will respond with error **AADSTS500133**.</span></span> <span data-ttu-id="9c4b3-268">В этом случае надстройке следует просто рекурсивно вызвать `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-268">In this case, the add-in should simply recursively call `getGraphData`.</span></span> <span data-ttu-id="9c4b3-269">Так как срок действия кэшированного маркера начальной загрузки истек, Office получит новый маркер из Azure AD.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-269">Since the cached bootstrap token is now expired, Office will get a new one from Azure AD.</span></span> <span data-ttu-id="9c4b3-270">Поэтому замените `TODO 8` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-270">So, replace `TODO 8` with the following markup:</span></span> 

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)       
    {
        getGraphData();
    }
    ```

1. <span data-ttu-id="9c4b3-271">Чтобы надстройка не вошла в бесконечный цикл вызовов `getGraphData`, она должна отслеживать число вызовов `getGraphData` и обеспечивать отсутствие повторных рекурсивных вызовов.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-271">To ensure that the add-in doesn't enter an infinite loop of calls to `getGraphData`, the add-in should keep track of how many times `getGraphData` has been called and be sure that is not called recursively called more than once.</span></span> <span data-ttu-id="9c4b3-272">Поэтому создайте переменную счетчика в области, которая является глобальной для функций `handleAADErrors` и `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-272">So, create a counter variable in a scope that is global to the `handleAADErrors` and `getGraphData` functions.</span></span> <span data-ttu-id="9c4b3-273">Подходящее место для глобальных переменных — сразу под вызовом метода `Office.onReady`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-273">A good place for global variables is just below the `Office.onReady` method call.</span></span>

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. <span data-ttu-id="9c4b3-274">Измените структуру `if` в методе `handleAADErrors`, чтобы он:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-274">Change the `if` structure in the `handleAADErrors` method so that it:</span></span>

    - <span data-ttu-id="9c4b3-275">увеличивал значение счетчика непосредственно перед вызовом `getGraphData`;</span><span class="sxs-lookup"><span data-stu-id="9c4b3-275">Increments the counter just before it calls `getGraphData`.</span></span>
    - <span data-ttu-id="9c4b3-276">выполнял тестирование, чтобы убедиться в отсутствии повторного вызова `getGraphData`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-276">Tests to ensure that `getGraphData` has not already been called a second time.</span></span> 

    <span data-ttu-id="9c4b3-277">Таким образом, окончательная версия структуры `if` должна выглядеть примерно так:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-277">So the final version of the `if` structure should look like the following:</span></span>

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="9c4b3-278">Замените `TODO 9` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-278">Replace `TODO 9` with the following.</span></span> 

    ```javascript
    else {                
        dialogFallback();
    }
    ```

1. <span data-ttu-id="9c4b3-279">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-279">Save and close the file.</span></span>

### <a name="get-the-data-and-add-it-to-the-office-document"></a><span data-ttu-id="9c4b3-280">Получение данных и их добавление в документ Office</span><span class="sxs-lookup"><span data-stu-id="9c4b3-280">Get the data and add it to the Office document</span></span>

1. <span data-ttu-id="9c4b3-281">Создайте в папке `public\javascripts` файл под названием `data.js`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-281">In the `public\javascripts` folder, create a new file named `data.js`, and paste the following code:</span></span>

1. <span data-ttu-id="9c4b3-282">Добавьте указанную ниже функцию в файл.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-282">Add the following function to the file.</span></span> <span data-ttu-id="9c4b3-283">Это функция, вызываемая функцией `getGraphData` при получении маркера доступа к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-283">This is the function that is called by the `getGraphData` function when it has acquired an access token to Microsoft Graph.</span></span> 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. <span data-ttu-id="9c4b3-284">Замените `TODO 10` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-284">Replace `TODO 10` with the following.</span></span> <span data-ttu-id="9c4b3-285">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-285">About this code, note:</span></span> 

    - <span data-ttu-id="9c4b3-286">Этот объект является параметром метода `$.ajax`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-286">This object is the parameter to the `$.ajax` method.</span></span>
    - <span data-ttu-id="9c4b3-287">`/getuserdata` — это экспресс-маршрут на сервере надстройки, создаваемый на более позднем шаге.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-287">The `/getuserdata` is an Express route on the add-in's server that you create in a later step.</span></span> <span data-ttu-id="9c4b3-288">Он вызывает конечную точку Microsoft Graph и добавляет маркер доступа в этот вызов.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-288">It will call a Microsoft Graph endpoint and include the access token in its call.</span></span> 

    ```javascript
    {
        type: "GET", 
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. <span data-ttu-id="9c4b3-289">Замените `TODO11` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-289">Replace `TODO11` with the following.</span></span> <span data-ttu-id="9c4b3-290">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-290">About this code, note:</span></span>

    - <span data-ttu-id="9c4b3-291">`writeFileNamesToOfficeDocument` вставляет данные из Graph в документ Office.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-291">The `writeFileNamesToOfficeDocument` will insert the data from Graph into the Office document.</span></span> <span data-ttu-id="9c4b3-292">Он определен в файле `public\javascripts\document.js`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-292">The `public\javascripts\document.js` method is defined in the src\auth.ts file.</span></span> 
    - <span data-ttu-id="9c4b3-293">Если `writeFileNamesToOfficeDocument` возвращает ошибку, она начнется с сообщения "Не удалось добавить имена файлов в документ".</span><span class="sxs-lookup"><span data-stu-id="9c4b3-293">If `writeFileNamesToOfficeDocument` returns an error, it will begin with "Unable to add filenames to document."</span></span>

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () { 
        showMessage("Your data has been added to the document."); 
    })
    .catch(function (error) {        
        showMessage(error);
    });
    ```

1. <span data-ttu-id="9c4b3-294">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-294">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="9c4b3-295">Код на стороне сервера</span><span class="sxs-lookup"><span data-stu-id="9c4b3-295">Code the server-side</span></span>

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a><span data-ttu-id="9c4b3-296">Создание маршрутизатора проверки подлинности и логики обмена маркерами</span><span class="sxs-lookup"><span data-stu-id="9c4b3-296">Create the auth router and the token exchange logic</span></span>

1. <span data-ttu-id="9c4b3-297">Откройте файл `routes\authRoute.js` и добавьте следующую функцию маршрутизации непосредственно под операторами `require` и над оператором `module.exports`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-297">Open the file `routes\authRoute.js` and add the following route function just below the `require` statements and above the `module.exports` statement.</span></span> <span data-ttu-id="9c4b3-298">Обратите внимание, что параметр URL-адреса `router.get` имеет значение '/'.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-298">Note that the URL parameter of `router.get` is '/'.</span></span> <span data-ttu-id="9c4b3-299">Так как этот маршрут определен в маршрутизаторе, обрабатывающем все HTTP-запросы для URL-адреса '/auth', этот маршрут эффективно обрабатывает все запросы для '/auth'.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-299">Since this route is being defined in a router that will handle all HTTP Requests for the URL '/auth', this route effectively handles all requests for '/auth'.</span></span> <span data-ttu-id="9c4b3-300">Клиентская функция `getGraphToken`, созданная ранее, вызывает этот маршрут.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-300">The client-side `getGraphToken` function that you created earlier calls this route.</span></span>  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exhange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. <span data-ttu-id="9c4b3-301">Замените `TODO 12` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-301">Replace `TODO 12` with the following code.</span></span>

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. <span data-ttu-id="9c4b3-302">Замените `TODO 13` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-302">Replace `TODO 13` with the following code.</span></span> <span data-ttu-id="9c4b3-303">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-303">About this code, note:</span></span> 

    - <span data-ttu-id="9c4b3-304">Это начало длинного блока `else`, но закрывающая скобка `}` не находится в конце, так как будет добавлен дополнительный код.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-304">This is the beginning of a long `else` block, but the closing `}` is not at the end yet because you will be adding more code to it.</span></span> 
    - <span data-ttu-id="9c4b3-305">Строка `authorization` — "носитель", за которым следует маркер начальной загрузки. Поэтому первая строка блока `else` присваивает маркер для `jwt`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-305">The `authorization` string is "Bearer " followed by the bootstrap token, so the first line of the `else` block is assigning the token to the `jwt`.</span></span> <span data-ttu-id="9c4b3-306">(JWT означает "веб-маркер JSON".)</span><span class="sxs-lookup"><span data-stu-id="9c4b3-306">("JWT" stands for "JSON Web Token".)</span></span>
    - <span data-ttu-id="9c4b3-307">Два значения `process.env.*` — это константы, назначаемые при настройке надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-307">The two `process.env.*` values are the constants that you assigned when you configured the add-in.</span></span> 
    - <span data-ttu-id="9c4b3-308">Параметру формы `requested_token_use` присвоено значение 'on_behalf_of'.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-308">The `requested_token_use` form parameter is set to 'on_behalf_of'.</span></span> <span data-ttu-id="9c4b3-309">Это указывает Azure AD, что надстройка запрашивает маркер доступа к Microsoft Graph, используя поток "от имени".</span><span class="sxs-lookup"><span data-stu-id="9c4b3-309">This tells Azure AD that the add-in is requesting an access token to Microsoft Graph using the On-Behalf-Of Flow.</span></span> <span data-ttu-id="9c4b3-310">Azure ответит проверкой того, что маркер начальной загрузки, назначенный параметру формы `assertion`, содержит свойство `scp` со значением `access-as-user`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-310">Azure will respond by validating that the bootstrap token, which is assigned to `assertion` form parameter, has a `scp` property that is set to `access-as-user`.</span></span>
    - <span data-ttu-id="9c4b3-311">Параметру формы `scope` присвоено значение 'Files.Read.All', что является единственной областью Microsoft Graph, требующейся надстройке.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-311">The `scope` form parameter is set to 'Files.Read.All' which is the only Microsoft Graph scope that the add-in needs.</span></span>

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. <span data-ttu-id="9c4b3-312">Замените `TODO 14` приведенным ниже кодом, дополняющим блок `else`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-312">Replace `TODO 14` with the following code, which completes the `else` block.</span></span> <span data-ttu-id="9c4b3-313">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-313">About this code, note:</span></span>

    - <span data-ttu-id="9c4b3-314">Константе `tenant` присвоено значение 'common', так как вы сделали надстройку мультитенатной при ее регистрации в Azure AD; в частности, когда назначили параметру **Поддерживаемые типы учетных записей** значение **Учетные записи в любом каталоге организации и персональные учетные записи Майкрософт (например, Skype, Xbox, Outlook.com)**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-314">The const `tenant` is set to 'common' because you configured the add-in as multitenant when you registered it with Azure AD; specifically when you set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span> <span data-ttu-id="9c4b3-315">Если вы решили поддерживать учетные записи только в том клиенте Office 365, где зарегистрирована надстройка, в этом коде `tenant` будет указан идентификатор GUID клиента.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-315">If you had instead chosen to support only accounts in the same Office 365 tenancy where the add-in is registered, then in this code `tenant` would be set to the GUID of the tenant.</span></span> 
    - <span data-ttu-id="9c4b3-316">Если при запросе POST не возникает ошибка, ответ от Azure AD преобразуется в формат JSON и отправляется клиенту.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-316">If the POST request does not error, then the response from Azure AD is converted to JSON and sent to the client.</span></span> <span data-ttu-id="9c4b3-317">Этот объект JSON содержит свойство `access_token`, которому служба Azure AD назначила маркер доступа в Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-317">This JSON object has an `access_token` property to which Azure AD has assigned the access token to Microsoft Graph.</span></span>

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: form(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();
            
            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. <span data-ttu-id="9c4b3-318">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-318">Save and close the file.</span></span>

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a><span data-ttu-id="9c4b3-319">Создание маршрута для извлечения данных из Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="9c4b3-319">Create the route that will fetch the data from Microsoft Graph</span></span>

1. <span data-ttu-id="9c4b3-320">Откройте файл `app.js` в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-320">Open the Startup.cs file in the root of the project.</span></span> <span data-ttu-id="9c4b3-321">Сразу под маршрутом для '/dialog.html' добавьте следующий маршрут.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-321">Just below the route for '/dialog.html', add the following route.</span></span> <span data-ttu-id="9c4b3-322">Этот маршрут вызывается функцией `makeGraphApiCall`, созданной на предыдущем шаге.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-322">This route is called by the `makeGraphApiCall` function that you created in an earlier step.</span></span>

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. <span data-ttu-id="9c4b3-323">Замените `TODO 15` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-323">Replace `TODO 15` with the following.</span></span> <span data-ttu-id="9c4b3-324">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-324">About this code, note:</span></span>

    - <span data-ttu-id="9c4b3-325">Метод `makeGraphApiCall`, вызывающий этот маршрут, добавляет маркер доступа к Microsoft Graph в HTTP-запрос в качестве заголовка с именем access_token.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-325">The caller of this route, `makeGraphApiCall`, added the access token to Microsoft Graph to the HTTP Request as a header named "access_token".</span></span>
    - <span data-ttu-id="9c4b3-326">Функция `getGraphData`определена в файле `msgraph-helper.js`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-326">The  method is defined in the src\auth.ts file.</span></span> <span data-ttu-id="9c4b3-327">(Эта функция отличается от клиентской функции `getGraphData`, определенной в файле `ssoAuthES6.js`.)</span><span class="sxs-lookup"><span data-stu-id="9c4b3-327">(This is not the same function as the client-side `getGraphData` function that you defined in the `ssoAuthES6.js` file.)</span></span>
    - <span data-ttu-id="9c4b3-328">Последний параметр для `queryParamsSegment` задается жестко.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-328">The last parameter, for `queryParamsSegment`, is hardcoded.</span></span> <span data-ttu-id="9c4b3-329">Если вы повторно используете этот код в рабочей надстройке и какая-либо часть `queryParamsSegment` получена из введенных пользователем данных, убедитесь, что он очищен и не может быть использован для атаки путем внедрения заголовка отклика.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-329">If you reuse this code in a production add-in and any part of `queryParamsSegment` comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.</span></span>
    - <span data-ttu-id="9c4b3-330">Код сводит к минимуму данные, которые должны поступать из Microsoft Graph, указывая только нужное свойство ("name") и только первые 10 имен папок или файлов.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-330">The code minimizes the data that must come from Microsoft Graph by specifying only the property we need ("name") and only the top 10 folder or file names.</span></span>

    ```javascript
    const graphToken = req.get('access_token');    
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. <span data-ttu-id="9c4b3-331">Замените `TODO 16` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-331">Replace `TODO 16` with the following.</span></span> <span data-ttu-id="9c4b3-332">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="9c4b3-332">About this code, note:</span></span>

    - <span data-ttu-id="9c4b3-333">Если Microsoft Graph возвращает ошибку (например, недопустимый или истекший маркер), возвращаемый объект будет содержать свойство кода со значением состояния HTTP (например, 401).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-333">If Microsoft Graph returns an error, such as invalid or expired token, there will be a code property in the returned object set to a HTTP status (e.g., 401).</span></span> <span data-ttu-id="9c4b3-334">Код передает ошибку клиенту.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-334">The code relays the error to the client.</span></span> <span data-ttu-id="9c4b3-335">Она перехватывается обратным вызовом `.fail` метода `makeGraphApiCall`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-335">It will be caught in the `.fail` callback of `makeGraphApiCall`.</span></span>
    - <span data-ttu-id="9c4b3-336">Данные Microsoft Graph включают метаданные OData и теги eTag, не требующиеся надстройке, поэтому код создает новый массив, содержащий только имена файлов для отправки клиенту.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-336">Microsoft Graph data includes OData metadata and eTags that the add-in does not need, so the code constructs a new array containing only the file names to send to the client.</span></span>

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. <span data-ttu-id="9c4b3-337">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-337">Save and close the file.</span></span>

## <a name="run-the-project"></a><span data-ttu-id="9c4b3-338">Запуск проекта</span><span class="sxs-lookup"><span data-stu-id="9c4b3-338">Run the project</span></span>

1. <span data-ttu-id="9c4b3-339">Убедитесь в наличии нескольких файлов в OneDrive, чтобы можно было проверить результаты.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-339">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="9c4b3-340">Откройте командную строку в корне папки `\Complete`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-340">Open a command prompt in the root of the `\Complete` folder.</span></span> 

1. <span data-ttu-id="9c4b3-341">Выполните команду `npm start`.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-341">Run the command  `npm start`.</span></span> 

1. <span data-ttu-id="9c4b3-342">Вам потребуется загрузить неопубликованную надстройку в приложение Office (Excel, Word или PowerPoint), чтобы протестировать ее.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-342">You need to sideload the add-in into an Office application (Excel, Word, or PowerPoint) to test it.</span></span> <span data-ttu-id="9c4b3-343">Инструкции зависят от вашей платформы.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-343">The instructions depend on your platform.</span></span> <span data-ttu-id="9c4b3-344">Ссылки на инструкции доступны в разделе [Загрузка неопубликованной надстройки Office для тестирования](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-344">There are links to instructions at [Sideload an Office Add-in for Testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).</span></span>

1. <span data-ttu-id="9c4b3-345">В приложении Office на вкладке ленты **Главная** нажмите кнопку **Показать надстройку** в группе **Единый вход Node.js**, чтобы открыть надстройку области задач.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-345">In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.</span></span>

1. <span data-ttu-id="9c4b3-346">Нажмите кнопку **Получить имена файлов OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-346">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="9c4b3-347">Если вы выполнили вход в Office с помощью рабочей или учебной (Office 365) учетной записи либо учетной записи Майкрософт и единый вход работает надлежащим образом, первые 10 имен файлов и папок из OneDrive для бизнеса вставляются в документ.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-347">If you are logged into Office with either a Work or School (Office 365) account or Microsoft Account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are inserted into the document.</span></span> <span data-ttu-id="9c4b3-348">(В первый раз это может занять до 15 секунд.) Если вы не выполнили вход или используете сценарий, не поддерживающий единый вход, или единый вход не работает по какой-то причине, появится запрос на вход.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-348">(It may take as much as 15 seconds the first time.) If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in.</span></span> <span data-ttu-id="9c4b3-349">После входа в систему отобразятся имена файлов и папок.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-349">After you log in, the file and folder names appear.</span></span>

> [!NOTE]
> <span data-ttu-id="9c4b3-350">Если вы ранее выполняли вход в Office с использованием другого идентификатора и все еще не закрыли некоторые из открытых тогда приложений Office, Office может не сменить идентификатор (даже если кажется, что это сделано).</span><span class="sxs-lookup"><span data-stu-id="9c4b3-350">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="9c4b3-351">Если это произойдет, возможен сбой при вызове Microsoft Graph или возврат данных для другого идентификатора.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-351">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="9c4b3-352">Чтобы избежать этого, *закройте все приложения Office*, прежде чем нажимать кнопку **Получить имена файлов OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="9c4b3-352">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>
