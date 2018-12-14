---
title: Создание надстройки Office на платформе Node.js с использованием единого входа
description: ''
ms.date: 12/07/2018
ms.openlocfilehash: 793d68dd3f1794c997a85bd5be682037aecca89f
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270994"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="27cf7-102">Создание надстройки Office на платформе Node.js с использованием единого входа (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="27cf7-102">Create a Node.js Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="27cf7-p101">Ваша веб-надстройка Office может использовать процедуру входа в Office для авторизации пользователей в надстройке и Microsoft Graph. При этом им не потребуется входить повторно. Общие сведения см. в статье [Включение единого входа в надстройке Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="27cf7-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="27cf7-105">Из этой статьи вы узнаете, как включить единый вход в надстройке, созданной с помощью Node.js и Express.</span><span class="sxs-lookup"><span data-stu-id="27cf7-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span> 

> [!NOTE]
> <span data-ttu-id="27cf7-106">Аналогичная статья, посвященная надстройке на основе ASP.NET, — [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md).</span><span class="sxs-lookup"><span data-stu-id="27cf7-106">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="27cf7-107">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="27cf7-107">Prerequisites</span></span>

* <span data-ttu-id="27cf7-108">[Node и npm](https://nodejs.org/en/) версии 6.9.4 или более поздней.</span><span class="sxs-lookup"><span data-stu-id="27cf7-108">[Node and npm](https://nodejs.org/en/), version 6.9.4 or later</span></span>

* <span data-ttu-id="27cf7-109">[Git Bash](https://git-scm.com/downloads) (или другой клиент git).</span><span class="sxs-lookup"><span data-stu-id="27cf7-109">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="27cf7-110">TypeScript версии 2.2.2 или более поздней.</span><span class="sxs-lookup"><span data-stu-id="27cf7-110">TypeScript version 2.2.2 or later</span></span>

* <span data-ttu-id="27cf7-111">Office 2016 версии 1708 (сборка 8424.nnnn) или более поздней (версия подписки на Office 365, иногда называемая "нажми и работай").</span><span class="sxs-lookup"><span data-stu-id="27cf7-111">Office 2016, Version 1708, build 8424.nnnn or later (the Office 365 subscription version, sometimes called “Click to Run”)</span></span>

  <span data-ttu-id="27cf7-p102">Для получения этой версии может понадобиться быть участником программы предварительной оценки Office. Дополнительные сведения см. в статье [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="27cf7-p102">You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="27cf7-114">Настройка начального проекта</span><span class="sxs-lookup"><span data-stu-id="27cf7-114">Set up the starter project</span></span>

1. <span data-ttu-id="27cf7-115">Клонируйте или скачайте репозиторий [Office-Add-in-NodeJS-SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span><span class="sxs-lookup"><span data-stu-id="27cf7-115">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="27cf7-116">Существует три версии примера.</span><span class="sxs-lookup"><span data-stu-id="27cf7-116">There are three versions of the sample:</span></span>  
    > * <span data-ttu-id="27cf7-p103">В папке **Before** находится начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы. В последующих разделах этой статьи рассматривается доработка проекта.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span> 
    > * <span data-ttu-id="27cf7-p104">Версия примера в папке **Completed** идентична надстройке, которую вы бы создали, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью. Чтобы использовать готовую версию, просто выполните действия, описанные в этой статье, но замените папку Before на папку Completed и пропустите разделы **Код на стороне клиента** и **Код на стороне сервера**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p104">The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="27cf7-122">Версия в папке **Completed Multitenant** — готовый пример, который поддерживает мультитенантность.</span><span class="sxs-lookup"><span data-stu-id="27cf7-122">The **Completed Multitenant** version is a completed sample that supports multitenancy.</span></span> <span data-ttu-id="27cf7-123">Изучите этот пример, если вы намерены поддерживать учетные записи Майкрософт с разных доменов с единым входом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-123">Explore this sample if you intend to support Microsoft accounts from different domains with SSO.</span></span>
    >
    > <span data-ttu-id="27cf7-124">_Вне зависимости от используемой версии вам понадобится сделать доверенным сертификат для localhost. См. примечание "ВАЖНО!" в файле сведений о репозитории._</span><span class="sxs-lookup"><span data-stu-id="27cf7-124">_Regardless of which version you use, you will need to trust a certificate for the localhost. See the "IMPORTANT" note in the Readme of the repo._</span></span>

2. <span data-ttu-id="27cf7-125">Откройте консоль Git bash в папке **Before**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-125">Open a Git bash console in the **Before** folder.</span></span>

3. <span data-ttu-id="27cf7-126">Введите в консоли команду `npm install`, чтобы установить все зависимости, указанные в файле package.json.</span><span class="sxs-lookup"><span data-stu-id="27cf7-126">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

4. <span data-ttu-id="27cf7-127">Введите в консоли команду `npm run build `, чтобы выполнить сборку проекта.</span><span class="sxs-lookup"><span data-stu-id="27cf7-127">Enter `npm run build ` in the console to build the project.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="27cf7-p106">Могут возникать ошибки сборки с сообщениями о том, что некоторые переменные объявлены, но не используются. Игнорируйте эти ошибки. Они возникают из-за того, что в исходной версии примера отсутствует код, который будет добавлен позже.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p106">You may see some build errors saying that some variables are declared but not used. Ignore these errors. They are a side effect of the fact that the "Before" version of the sample is missing some code that will be added later.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="27cf7-131">Регистрация надстройки в конечной точке Azure AD версии 2.0</span><span class="sxs-lookup"><span data-stu-id="27cf7-131">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="27cf7-132">Следующие инструкции содержат общую информацию, поэтому их можно использовать в нескольких местах.</span><span class="sxs-lookup"><span data-stu-id="27cf7-132">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="27cf7-133">В рамках этой статьи сделайте вот что:</span><span class="sxs-lookup"><span data-stu-id="27cf7-133">For this article do the following:</span></span>
- <span data-ttu-id="27cf7-134">Замените заполнитель **$ADD-IN-NAME$** на `“Office-Add-in-NodeJS-SSO`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-134">Replace the placeholder **$ADD-IN-NAME$** with `“Office-Add-in-NodeJS-SSO`.</span></span>
- <span data-ttu-id="27cf7-135">Замените заполнитель **$FQDN-WITHOUT-PROTOCOL$** на `localhost:3000`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-135">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:3000`.</span></span>
- <span data-ttu-id="27cf7-136">Указывая разрешения в диалоговом окне **Выбор разрешений**, установите флажки для приведенных ниже разрешений.</span><span class="sxs-lookup"><span data-stu-id="27cf7-136">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="27cf7-137">Для самой надстройки требуется только первое разрешение, но разрешение `profile` необходимо, чтобы ведущее приложение Office получило маркер для веб-приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="27cf7-137">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
    * <span data-ttu-id="27cf7-138">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="27cf7-138">Files.Read.All</span></span>
    * <span data-ttu-id="27cf7-139">profile</span><span class="sxs-lookup"><span data-stu-id="27cf7-139">profile</span></span>

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="27cf7-140">Предоставление надстройке разрешений администратора</span><span class="sxs-lookup"><span data-stu-id="27cf7-140">Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="27cf7-141">Настройка надстройки</span><span class="sxs-lookup"><span data-stu-id="27cf7-141">Configure the add-in</span></span>

1. <span data-ttu-id="27cf7-p109">В редакторе кода откройте файл src\server.ts. В начале этого файла есть вызов конструктора класса `AuthModule`. У конструктора есть строковые параметры, которым необходимо назначить значения.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p109">In your code editor, open the src\server.ts file. Near the top there is a call to a constructor of an `AuthModule` class. There are some string parameters in the constructor to which you need to assign values.</span></span>

2. <span data-ttu-id="27cf7-145">В свойстве `client_id` замените заполнитель `{client GUID}` на идентификатор приложения, сохраненный во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="27cf7-145">For the `client_id` property, replace the placeholder `{client GUID}` with the application ID that you saved when you registered the add-in.</span></span> <span data-ttu-id="27cf7-146">В результате должен остаться только GUID в одиночных кавычках.</span><span class="sxs-lookup"><span data-stu-id="27cf7-146">When you are done, there should just be a GUID in single quotation marks.</span></span> <span data-ttu-id="27cf7-147">Значение не должно содержать символов "{}".</span><span class="sxs-lookup"><span data-stu-id="27cf7-147">There should not be any "{}" characters.</span></span>

3. <span data-ttu-id="27cf7-148">В свойстве `client_secret` замените заполнитель `{client secret}` на секрет приложения, сохраненный во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="27cf7-148">For the `client_secret` property, replace the placeholder `{client secret}` with the application secret that you saved when you registered the add-in.</span></span>

4. <span data-ttu-id="27cf7-p111">В свойстве `audience` замените заполнитель `{audience GUID}` на идентификатор приложения, сохраненный во время регистрации надстройки. (Это то же значение, которое вы назначили свойству `client_id`.)</span><span class="sxs-lookup"><span data-stu-id="27cf7-p111">For the `audience` property, replace the placeholder `{audience GUID}` with the application ID that you saved when you registered the add-in. (The very same value that you assigned to the `client_id` property.)</span></span>
  
3. <span data-ttu-id="27cf7-151">В строке, назначенной свойству `issuer`, вы увидите заполнитель *{O365 tenant GUID}*.</span><span class="sxs-lookup"><span data-stu-id="27cf7-151">In the string assigned to the `issuer` property, you will see the placeholder *{O365 tenant GUID}*.</span></span> <span data-ttu-id="27cf7-152">Замените его идентификатором клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="27cf7-152">Replace this with the Office 365 tenancy ID.</span></span> <span data-ttu-id="27cf7-153">Чтобы получить его, используйте один из способов, описанных в статье [Поиск идентификатора клиента Office 365](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id).</span><span class="sxs-lookup"><span data-stu-id="27cf7-153">Use one of the methods in [Find your Office 365 tenant ID](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span> <span data-ttu-id="27cf7-154">В результате значение свойства `issuer` должно выглядеть примерно так:</span><span class="sxs-lookup"><span data-stu-id="27cf7-154">When you are done, the `issuer` property value should look something like this:</span></span>

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. <span data-ttu-id="27cf7-155">Оставьте остальные параметры конструктора `AuthModule` без изменений.</span><span class="sxs-lookup"><span data-stu-id="27cf7-155">Leave the other parameters in the `AuthModule` constructor unchanged.</span></span> <span data-ttu-id="27cf7-156">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="27cf7-156">Save and close the file.</span></span>

1. <span data-ttu-id="27cf7-157">В корневой папке проекта откройте файл манифеста Office-Add-in-NodeJS-SSO.xml.</span><span class="sxs-lookup"><span data-stu-id="27cf7-157">In the root of the project, open the add-in manifest file “Office-Add-in-NodeJS-SSO.xml”.</span></span>

1. <span data-ttu-id="27cf7-158">Прокрутите вниз до конца файла.</span><span class="sxs-lookup"><span data-stu-id="27cf7-158">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="27cf7-159">Над последним тегом `</VersionOverrides>` вы найдете следующую разметку:</span><span class="sxs-lookup"><span data-stu-id="27cf7-159">Just above the end `</VersionOverrides>` tag, you will find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="27cf7-160">Замените заполнитель {application_GUID here} *в обоих местах* разметки идентификатором приложения, скопированным при регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="27cf7-160">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="27cf7-161">(Символы "{}" не входят в состав идентификатора, поэтому их не нужно вставлять.) Это тот же идентификатор, который использовался для ClientID и Audience в файле web.config.</span><span class="sxs-lookup"><span data-stu-id="27cf7-161">(The "{}" are not part of the ID, so don't include them.) This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="27cf7-162">Значение **Resource** представляет собой **URI идентификатора приложения**, который вы задали, когда добавляли платформу веб-API при регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="27cf7-162">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="27cf7-163">Раздел **Scopes** используется для создания диалогового окна согласия, только если надстройка продается в AppSource.</span><span class="sxs-lookup"><span data-stu-id="27cf7-163">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="27cf7-164">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="27cf7-164">Save and close the file.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="27cf7-165">Код на стороне клиента</span><span class="sxs-lookup"><span data-stu-id="27cf7-165">Code the client side</span></span>

1. <span data-ttu-id="27cf7-p115">Откройте файл program.js в папке **public**. В нем уже есть следующий код:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p115">Open the program.js file in the **public** folder. It already has some code in it:</span></span>

    * <span data-ttu-id="27cf7-168">Назначение методу `Office.initialize`, которое, в свою очередь, назначает обработчик события для нажатия кнопки `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-168">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="27cf7-169">Метод `showResult` для отображения сообщения об ошибке (или данных, возвращаемых из Microsoft Graph) в нижней части области задач.</span><span class="sxs-lookup"><span data-stu-id="27cf7-169">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="27cf7-170">Метод `logErrors` для регистрации в консоли ошибок, которые не предназначены для пользователя.</span><span class="sxs-lookup"><span data-stu-id="27cf7-170">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

11. <span data-ttu-id="27cf7-p116">После назначения для метода `Office.initialize` добавьте приведенный ниже код. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p116">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="27cf7-173">При обработке ошибок в надстройке иногда автоматически выполняется еще одна попытка получить маркер доступа с помощью другого набора параметров.</span><span class="sxs-lookup"><span data-stu-id="27cf7-173">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="27cf7-174">Переменная счетчика `timesGetOneDriveFilesHasRun`, переменные флага `triedWithoutForceConsent` и `timesMSGraphErrorReceived` используются, чтобы для пользователя не повторялись циклически неудачные попытки получить маркер.</span><span class="sxs-lookup"><span data-stu-id="27cf7-174">The counter variable `timesGetOneDriveFilesHasRun`, and the flag variables `triedWithoutForceConsent` and `timesMSGraphErrorReceived` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span> 
    * <span data-ttu-id="27cf7-p118">Метод `getDataWithToken` создается на следующем шаге. Обратите внимание на то, что он присваивает параметру `forceConsent` значение `false`. Дополнительные сведения см. в описании следующего шага.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p118">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. <span data-ttu-id="27cf7-p119">Под методом `getOneDriveFiles` добавьте приведенный ниже код. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p119">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="27cf7-179">[getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) — это новый API в Office.js, позволяющий надстройке запрашивать у ведущего приложения Office (Excel, PowerPoint, Word и т. д.) маркер доступа к надстройке (для пользователя, выполнившего вход в Office).</span><span class="sxs-lookup"><span data-stu-id="27cf7-179">The [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office).</span></span> <span data-ttu-id="27cf7-180">Ведущее приложение Office, в свою очередь, запрашивает маркер у конечной точки Azure AD версии 2.0.</span><span class="sxs-lookup"><span data-stu-id="27cf7-180">The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token.</span></span> <span data-ttu-id="27cf7-181">Так как вы предварительно авторизовали ведущее приложение Office для надстройки во время ее регистрации, Azure AD отправит токен.</span><span class="sxs-lookup"><span data-stu-id="27cf7-181">Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="27cf7-182">Если вход в Office не выполнен, ведущее приложение Office предложит пользователю войти.</span><span class="sxs-lookup"><span data-stu-id="27cf7-182">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="27cf7-183">Параметр настроек задает для `forceConsent` значение `false`, поэтому пользователю не будет предлагаться разрешить ведущему приложению Office доступ к надстройке при каждом ее использовании.</span><span class="sxs-lookup"><span data-stu-id="27cf7-183">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in.</span></span> <span data-ttu-id="27cf7-184">При первом запуске надстройки вызов `getAccessTokenAsync` не будет выполнен, но логика обработки ошибок, которую вы добавите на следующем этапе, автоматически выполнит повторный вызов, при этом параметру `forceConsent` будет задано значение `true`, и пользователю будет предложено согласиться. Такая процедура выполняется только в первый раз.</span><span class="sxs-lookup"><span data-stu-id="27cf7-184">The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="27cf7-185">Вы создадите метод `handleClientSideErrors` позже.</span><span class="sxs-lookup"><span data-stu-id="27cf7-185">You will create the `handleClientSideErrors` method in a later step.</span></span>

    ```javascript
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. <span data-ttu-id="27cf7-p122">Замените строку TODO1 на приведенные ниже строки. Метод `getData` и серверный маршрут /api/values создаются позже. Для конечной точки используется относительный URL-адрес, так как она должна размещаться на том же домене, что и надстройка.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p122">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="27cf7-p123">Под методом `getOneDriveFiles` добавьте приведенный ниже код. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p123">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="27cf7-p124">Этот метод вызывает указанную конечную точку веб-API и передает ей тот же маркер доступа, который ведущее приложение Office использовало для доступа к надстройке. На стороне сервера этот маркер доступа будет использоваться в потоке "от имени" для получения маркера доступа к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p124">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="27cf7-193">Вы создадите метод `handleServerSideErrors` позже.</span><span class="sxs-lookup"><span data-stu-id="27cf7-193">You will create the `handleServerSideErrors` method in a later step.</span></span>

    ```javascript
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        }); 
    }
    ```

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="27cf7-194">Создание методов обработки ошибок</span><span class="sxs-lookup"><span data-stu-id="27cf7-194">Create the error-handling methods</span></span>

1. <span data-ttu-id="27cf7-195">Под методом `getData` добавьте приведенный ниже метод.</span><span class="sxs-lookup"><span data-stu-id="27cf7-195">Below the `getData` method, add the following method.</span></span> <span data-ttu-id="27cf7-196">Этот метод будет обрабатывать ошибки в клиенте надстройки, когда ведущее приложение Office не сможет получить маркер доступа к веб-службе надстройки.</span><span class="sxs-lookup"><span data-stu-id="27cf7-196">This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service.</span></span> <span data-ttu-id="27cf7-197">Сообщения о таких ошибках содержат код ошибки, поэтому данный метод различает их с помощью оператора `switch`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-197">These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Microsoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user triggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="27cf7-198">Замените `TODO2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-198">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="27cf7-199">Ошибка 13001 возникает, если пользователь не выполнил вход или без отклика отменил запрос на предоставление 2-го фактора проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="27cf7-199">Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor.</span></span> <span data-ttu-id="27cf7-200">В обоих случаях код повторно выполняет метод `getDataWithToken` и задает параметр для принудительного запрашивания входа.</span><span class="sxs-lookup"><span data-stu-id="27cf7-200">In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="27cf7-201">Замените `TODO3` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-201">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="27cf7-202">Ошибка 13002 возникает, когда вход или предоставление разрешений прерывается.</span><span class="sxs-lookup"><span data-stu-id="27cf7-202">Error 13002 occurs when user's sign-in or consent was aborted.</span></span> <span data-ttu-id="27cf7-203">Попросите пользователя повторить попытку, но не более одного раза.</span><span class="sxs-lookup"><span data-stu-id="27cf7-203">Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. <span data-ttu-id="27cf7-204">Замените `TODO4` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-204">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="27cf7-205">Ошибка 13003 возникает, когда пользователь входит под учетной записью, отличной от рабочей, учебной или личной учетной записи Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="27cf7-205">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft Account.</span></span> <span data-ttu-id="27cf7-206">Попросите пользователя выйти, а затем войти с помощью учетной записи поддерживаемого типа.</span><span class="sxs-lookup"><span data-stu-id="27cf7-206">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > <span data-ttu-id="27cf7-207">Ошибки 13004 и 13005 не обрабатываются при использовании этого метода, так как они должны возникать только на стадии разработки.</span><span class="sxs-lookup"><span data-stu-id="27cf7-207">Errors 13004 and 13005 are not handled in this method because they should only occur in development.</span></span> <span data-ttu-id="27cf7-208">Их невозможно исправить с помощью кода среды выполнения, поэтому нет смысла сообщать о них пользователю.</span><span class="sxs-lookup"><span data-stu-id="27cf7-208">They cannot be fixed by runtime code and there would be no point in reporting them to an end user.</span></span>

1. <span data-ttu-id="27cf7-p130">Замените `TODO5` приведенным ниже кодом. Ошибка 13006 возникает, если происходит неопределенная ошибка ведущего приложения Office, которая может свидетельствовать о его нестабильном состоянии. Попросите пользователя перезапустить Office.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p130">Replace `TODO5` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. <span data-ttu-id="27cf7-212">Замените `TODO6` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-212">Replace `TODO6` with the following code.</span></span> <span data-ttu-id="27cf7-213">Ошибка 13007 возникает, когда нарушается взаимодействие ведущего приложения Office с AAD, из-за чего это приложение не может получить маркер доступа к веб-службе/приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="27cf7-213">Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application.</span></span> <span data-ttu-id="27cf7-214">Это может быть из-за временного сбоя сети.</span><span class="sxs-lookup"><span data-stu-id="27cf7-214">This may be a temporary network issue.</span></span> <span data-ttu-id="27cf7-215">Попросите пользователя повторить попытку позже.</span><span class="sxs-lookup"><span data-stu-id="27cf7-215">Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. <span data-ttu-id="27cf7-216">Замените `TODO7` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-216">Replace `TODO7` with the following code.</span></span> <span data-ttu-id="27cf7-217">Ошибка 13008 возникает, когда пользователь запускает операцию, которая вызывает `getAccessTokenAsync`, до завершения предыдущего вызова.</span><span class="sxs-lookup"><span data-stu-id="27cf7-217">Error 13008 occurs when the user triggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. <span data-ttu-id="27cf7-218">Замените `TODO8` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-218">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="27cf7-219">Ошибка 13009 возникает, если надстройка не поддерживает принудительное запрашивание разрешения, но выполняется вызов `getAccessTokenAsync` с установкой для параметра `forceConsent` значения `true`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-219">Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`.</span></span> <span data-ttu-id="27cf7-220">Обычно в таком случае код должен автоматически повторно запустить метод `getAccessTokenAsync` с параметром, имеющим значение `false`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-220">In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`.</span></span> <span data-ttu-id="27cf7-221">Но в некоторых случаях вызов метода с установкой для параметра `forceConsent` значения `true` сам по себе является автоматическим откликом на ошибку вызова метода с установкой для параметра значения `false`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-221">However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`.</span></span> <span data-ttu-id="27cf7-222">В этом случае код должен не повторять попытку, а предложить пользователю выйти и войти заново.</span><span class="sxs-lookup"><span data-stu-id="27cf7-222">In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. <span data-ttu-id="27cf7-223">Замените `TODO9` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-223">Replace `TODO9` with the following code.</span></span>

    ```javascript
    default:
        logError(result);
        break;
    ```  

1. <span data-ttu-id="27cf7-p134">Под методом `handleClientSideErrors` добавьте приведенный ниже метод. Этот метод обрабатывает ошибки в веб-службе надстройки при неправильном выполнении потока "от имени" или получении данных от Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p134">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Handle the case where AAD asks for an additional form of authentication.

        // TODO11: Handle the case where consent has not been granted, or has been revoked.

        // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO13: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. <span data-ttu-id="27cf7-p135">Замените `TODO10` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p135">Replace `TODO10` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="27cf7-p136">Существуют конфигурации Azure Active Directory, согласно которым пользователю необходимо предоставить дополнительные факторы проверки подлинности для доступа к некоторым целевым объектам Microsoft Graph (например, OneDrive), даже если пользователь может войти в Office, указав всего лишь пароль. В таком случае AAD отправит отклик, содержащий номер ошибки 50076 со свойством `Claims`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p136">There are configurations of Azure Active Directory in which the user is required to provide additional authentication factor(s) to access some Microsoft Graph targets (e.g., OneDrive), even if the user can sign on to Office with just a password. In that case, AAD will send a response, with error 50076, that has a `Claims` property.</span></span> 
    * <span data-ttu-id="27cf7-230">Ведущее приложение Office должно получить новый маркер со значением **Claims** в качестве параметра `authChallenge`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-230">The Office host should get a new token with the **Claims** value as the `authChallenge` option.</span></span> <span data-ttu-id="27cf7-231">Так AAD получит команду отобразить для пользователя запрос на прохождение всех форм проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="27cf7-231">This tells AAD to prompt the user for all required forms of authentication.</span></span> 

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. <span data-ttu-id="27cf7-p138">Замените `TODO11` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем шаге*. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p138">Replace `TODO11` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="27cf7-234">Ошибка 65001 означает, что доступ к Microsoft Graph не был предоставлен (или был отозван) для одного или нескольких разрешений.</span><span class="sxs-lookup"><span data-stu-id="27cf7-234">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span> 
    * <span data-ttu-id="27cf7-235">Надстройка должна получить новый маркер (параметру `forceConsent` должно быть задано значение `true`).</span><span class="sxs-lookup"><span data-stu-id="27cf7-235">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        getDataWithToken({ forceConsent: true });
    }
    ```

1. <span data-ttu-id="27cf7-p139">Замените `TODO12` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем шаге*. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p139">Replace `TODO12` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="27cf7-238">Ошибка 70011 означает, что запрошена недопустимая область (разрешение).</span><span class="sxs-lookup"><span data-stu-id="27cf7-238">Error 70011 means that an invalid scope (permission) has been requested.</span></span> <span data-ttu-id="27cf7-239">Надстройка должна сообщить об ошибке.</span><span class="sxs-lookup"><span data-stu-id="27cf7-239">The add-in should report the error.</span></span>
    * <span data-ttu-id="27cf7-240">Код регистрирует любую другую ошибку с номером ошибки AAD.</span><span class="sxs-lookup"><span data-stu-id="27cf7-240">The code logs any other error with an AAD error number.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. <span data-ttu-id="27cf7-p141">Замените `TODO13` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем шаге*. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p141">Replace `TODO13` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="27cf7-243">Код на стороне сервера, который вы создадите на более позднем этапе, отправит сообщение, заканчивающееся на `... expected access_as_user`, если область (разрешение) `access_as_user` будет отсутствовать в маркере доступа, отправляемом клиентом надстройки в AAD для использования в потоке "от имени".</span><span class="sxs-lookup"><span data-stu-id="27cf7-243">Server-side code that you create in a later step will send the message that ends with `... expected access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="27cf7-244">Надстройка должна сообщить об ошибке.</span><span class="sxs-lookup"><span data-stu-id="27cf7-244">The add-in should report the error.</span></span>

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. <span data-ttu-id="27cf7-p142">Замените `TODO14` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем шаге*. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p142">Replace `TODO14` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="27cf7-247">Маловероятно, чтобы в Microsoft Graph был отправлен недействительный маркер или маркер с истекшим сроком действия. Но если это произойдет, код на стороне сервера, который вы создадите на более позднем этапе, будет заканчиваться строкой `Microsoft Graph error`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-247">It is unlikely that an expired or invalid token will be sent to Microsoft Graph; but if it does happen, the server-side code that you will create in a later step will end with the string `Microsoft Graph error`.</span></span>
    * <span data-ttu-id="27cf7-248">В этом случае надстройка должна начать заново весь процесс проверки подлинности, сбросив счетчик `timesGetOneDriveFilesHasRun` и переменные флага `timesGetOneDriveFilesHasRun`, а затем повторно вызвать метод обработчика кнопок.</span><span class="sxs-lookup"><span data-stu-id="27cf7-248">In this case, the add-in should start the entire authentication process over by resetting the `timesGetOneDriveFilesHasRun` counter and `timesGetOneDriveFilesHasRun` flag variables, and then re-calling the button handler method.</span></span> <span data-ttu-id="27cf7-249">Но она должна сделать это только один раз.</span><span class="sxs-lookup"><span data-stu-id="27cf7-249">But it should do this only once.</span></span> <span data-ttu-id="27cf7-250">Если ситуация повторится, надстройка должна просто зарегистрировать ошибку.</span><span class="sxs-lookup"><span data-stu-id="27cf7-250">If it happens again, it should just log the error.</span></span>
    * <span data-ttu-id="27cf7-251">Код зарегистрирует ошибку, если она повторится два раза подряд.</span><span class="sxs-lookup"><span data-stu-id="27cf7-251">The code logs the error if it happens twice in succession.</span></span>

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }        
    }
    ```

1. <span data-ttu-id="27cf7-252">Замените `TODO15` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем этапе*.</span><span class="sxs-lookup"><span data-stu-id="27cf7-252">Replace `TODO15` with the following code *just below the last closing brace of the code you added in the previous step*.</span></span>

    ```javascript
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a><span data-ttu-id="27cf7-253">Код на стороне сервера</span><span class="sxs-lookup"><span data-stu-id="27cf7-253">Code the server side</span></span>

<span data-ttu-id="27cf7-254">На стороне сервера необходимо изменить два файла.</span><span class="sxs-lookup"><span data-stu-id="27cf7-254">There are two server-side files that need to be modified.</span></span> 
- <span data-ttu-id="27cf7-p144">Файл src\auth.js предоставляет вспомогательные функции авторизации. Он уже содержит универсальные элементы, используемые в различных потоках авторизации. Нам необходимо добавить в него функции, реализующие поток "от имени".</span><span class="sxs-lookup"><span data-stu-id="27cf7-p144">The src\auth.js provides authorization helper functions. It already has generic members that are used in a variety of authorization flows. We need to add functions to it that implement the "on behalf of" flow.</span></span>
- <span data-ttu-id="27cf7-p145">Файл src\server.js содержит базовые элементы, необходимые для запуска сервера и ПО промежуточного слоя express. Нам необходимо добавить в него функции, предоставляющие домашнюю страницу, и веб-API для получения данных Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p145">The src\server.js file has the basic members need to run a server and express middleware. We need to add functions to it that serve the home page and a Web API for obtaining Microsoft Graph data.</span></span>

### <a name="create-a-method-to-exchange-tokens"></a><span data-ttu-id="27cf7-260">Создание метода для обмена маркерами</span><span class="sxs-lookup"><span data-stu-id="27cf7-260">Create a method to exchange tokens</span></span>

1. <span data-ttu-id="27cf7-p146">Откройте файл \src\auth.ts. Добавьте приведенный ниже метод в класс `AuthModule`. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p146">Open the \src\auth.ts file. Add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="27cf7-p147">Параметр `jwt` — это маркер доступа к приложению. В потоке "от имени" он отправляется службе AAD в обмен на маркер доступа к ресурсу.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p147">The `jwt` parameter is the access token to the application. In the "on behalf of" flow, it is exchanged with AAD for an access token to the resource.</span></span>
    * <span data-ttu-id="27cf7-266">Параметр scopes содержит значение по умолчанию, но в этом примере его переопределяет код вызова.</span><span class="sxs-lookup"><span data-stu-id="27cf7-266">The scopes parameter has a default value, but in this sample it will be overridden by the calling code.</span></span>
    * <span data-ttu-id="27cf7-267">Указывать параметр resource не обязательно.</span><span class="sxs-lookup"><span data-stu-id="27cf7-267">The resource parameter is optional.</span></span> <span data-ttu-id="27cf7-268">Его не следует использовать, если [службой токенов безопасности (STS)](https://docs.microsoft.com/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) является конечная точка AAD версии 2.0.</span><span class="sxs-lookup"><span data-stu-id="27cf7-268">It should not be used when the [Secure Token Service (STS)](https://docs.microsoft.com/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) is the AAD V 2.0 endpoint.</span></span> <span data-ttu-id="27cf7-269">Конечная точка версии 2.0 получает ресурс из областей и возвращает ошибку, если ресурс отправлен в HTTP-запросе.</span><span class="sxs-lookup"><span data-stu-id="27cf7-269">The V 2.0 endpoint infers the resource from the scopes and it returns an error if a resource is sent in the HTTP Request.</span></span> 
    * <span data-ttu-id="27cf7-270">Выдача исключения в блоке `catch` *не* приводит к немедленной отправке текста "500 внутренняя ошибка сервера" клиенту.</span><span class="sxs-lookup"><span data-stu-id="27cf7-270">Throwing an exception in the `catch` block will *not* cause an immediate "500 Internal Server Error" to be sent to the client.</span></span> <span data-ttu-id="27cf7-271">Вызов кода в файле server.js захватывает данное исключение и преобразует его в сообщение об ошибке, отправляемое клиенту.</span><span class="sxs-lookup"><span data-stu-id="27cf7-271">Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

        ```typescript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the 
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                            + JSON.stringify(exception), 
                                            exception);
            }
        }
        ```

2. <span data-ttu-id="27cf7-p150">Замените `TODO3` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p150">Replace `TODO3` with the following code. About this code, note:</span></span>
    * <span data-ttu-id="27cf7-p151">Служба токенов безопасности, поддерживающая поток "от имени", ожидает определенные пары "ключ-значение" в тексте HTTP-запроса. Этот код конструирует объект, который станет текстом запроса.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p151">An STS that supports the "on behalf of" flow expects certain property/value pairs in the body of the HTTP request. This code constructs an object that will become the body of the request.</span></span> 
    * <span data-ttu-id="27cf7-276">Свойство ресурса добавляется в текст, только если методу был передан ресурс.</span><span class="sxs-lookup"><span data-stu-id="27cf7-276">A resource property is added to the body if, and only if, a resource was passed to the method.</span></span>

        ```typescript
        const v2Params = {
                client_id: this.clientId,
                client_secret: this.clientSecret,
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt,
                requested_token_use: 'on_behalf_of',
                scope: scopes.join(' ')
            };
            let finalParams = {};
            if (resource) {
                // In JavaScript we could just add the resource property to the v2Params
                // object, but that won't compile in TypeScript.
                let v1Params  = { resource: resource };  
                for(var key in v2Params) { v1Params[key] = v2Params[key]; }
                finalParams = v1Params;
            } else {
                finalParams = v2Params;
            } 
        ```

3. <span data-ttu-id="27cf7-277">Замените `TODO4` приведенным ниже кодом, который отправляет HTTP-запрос конечной точке маркера для службы токенов безопасности.</span><span class="sxs-lookup"><span data-stu-id="27cf7-277">Replace `TODO4` with the following code which sends the HTTP request to the token endpoint of the STS.</span></span>

    ```typescript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. <span data-ttu-id="27cf7-278">Замените `TODO5` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-278">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="27cf7-279">Обратите внимание на то, что выдача исключения *не* приводит к немедленной отправке текста "500 внутренняя ошибка сервера" клиенту.</span><span class="sxs-lookup"><span data-stu-id="27cf7-279">Note that throwing an exception will *not* cause an immediate "500 Internal Server Error" to be sent to the client.</span></span> <span data-ttu-id="27cf7-280">Вызов кода в файле server.js захватывает данное исключение и преобразует его в сообщение об ошибке, отправляемое клиенту.</span><span class="sxs-lookup"><span data-stu-id="27cf7-280">Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

    ```typescript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;                
    } 
    ```

5. <span data-ttu-id="27cf7-p153">Замените `TODO6` приведенным ниже кодом. Обратите внимание на то, что код сохраняет маркер доступа для ресурса и срок его действия, а не только возвращает его. В коде вызова можно обойтись без лишних вызовов службы токенов безопасности, повторно используя действительный маркер доступа к ресурсу. В следующем разделе показано, как это сделать.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p153">Replace `TODO6` with the following code. Note that the code persists the access token to the resource, and it's expiration time, in addition to returning it. Calling code can avoid unnecessary calls to the STS by reusing an unexpired access token to the resource. You'll see how to do that in the next section.</span></span>

    ```typescript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

6. <span data-ttu-id="27cf7-285">Сохраните файл, но не закрывайте его.</span><span class="sxs-lookup"><span data-stu-id="27cf7-285">Save the file, but don't close it.</span></span>

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a><span data-ttu-id="27cf7-286">Создание метода для доступа к ресурсу с помощью потока "от имени"</span><span class="sxs-lookup"><span data-stu-id="27cf7-286">Create a method to get access to the resource using the "on behalf of" flow</span></span>

1. <span data-ttu-id="27cf7-p154">В файле src/auth.ts добавьте метод под классом `AuthModule`. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p154">Still in src/auth.ts, add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="27cf7-289">Приведенные выше комментарии к параметрам метода `exchangeForToken` также применимы к параметрам данного метода.</span><span class="sxs-lookup"><span data-stu-id="27cf7-289">The comments above about the parameters to the the `exchangeForToken` method apply to the parameters of this method as well.</span></span>
    * <span data-ttu-id="27cf7-p155">Метод сначала проверяет постоянное хранилище на наличие действительного маркера доступа к ресурсу, срок действия которого не истечет через минуту. Он вызывает метод `exchangeForToken`, создание которого описано в предыдущем разделе, только если это необходимо.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p155">The method first checks the persistent storage for an access token to the resource that has not expired and is not going to expire in the next minute. It calls the `exchangeForToken` method you created in the last section only if it needs to.</span></span>

    ```typescript
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    } 
    ```

2. <span data-ttu-id="27cf7-292">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="27cf7-292">Save and close the file.</span></span>

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a><span data-ttu-id="27cf7-293">Создание конечных точек, предоставляющих домашнюю страницу и данные надстройки</span><span class="sxs-lookup"><span data-stu-id="27cf7-293">Create the endpoints that will serve the add-in's home page and data</span></span>

1. <span data-ttu-id="27cf7-294">Откройте файл src\server.ts.</span><span class="sxs-lookup"><span data-stu-id="27cf7-294">Open the src\server.ts file.</span></span> 

2. <span data-ttu-id="27cf7-p156">Добавьте приведенный ниже метод в конец файла. Этот метод будет предоставлять домашнюю страницу надстройки. В манифесте надстройки указан URL-адрес домашней страницы.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p156">Add the following method to the bottom of the file. This method will serve the add-in's home page. The add-in manifest specifies the home page URL.</span></span>

    ```typescript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. <span data-ttu-id="27cf7-p157">Добавьте приведенный ниже метод в конец файла. Этот метод будет обрабатывать все запросы к API `values`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p157">Add the following method to bottom of the file. This method will handle any requests for the `values` API.</span></span>
    ```typescript
    app.get('/api/values', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    })); 
    ```

4. <span data-ttu-id="27cf7-300">Замените `TODO7` приведенным ниже кодом, который проверяет маркер доступа, полученный от ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="27cf7-300">Replace `TODO7` with the following code which validates the access token received from the Office host application.</span></span> <span data-ttu-id="27cf7-301">Метод `verifyJWT` определен в файле src\auth.ts.</span><span class="sxs-lookup"><span data-stu-id="27cf7-301">The `verifyJWT` method is defined in the src\auth.ts file.</span></span> <span data-ttu-id="27cf7-302">Он всегда проверяет аудиторию и издателя.</span><span class="sxs-lookup"><span data-stu-id="27cf7-302">It always validates the audience and the issuer.</span></span> <span data-ttu-id="27cf7-303">С помощью необязательного параметра мы указываем на необходимость проверить, указана ли в маркере доступа область `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="27cf7-303">We use the optional parameter to specify that we also want it to verify that the scope in the access token is `access_as_user`.</span></span> <span data-ttu-id="27cf7-304">Это единственное разрешение для надстройки, необходимое пользователю и ведущему приложению Office, чтобы получить маркер доступа к Microsoft Graph с помощью потока "от имени".</span><span class="sxs-lookup"><span data-stu-id="27cf7-304">This is the only permission to the add-in that the user and the Office host need in order to get an access token to Microsoft Graph by means of the "on behalf" flow.</span></span> 

    ```typescript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

    > [!NOTE]
    > <span data-ttu-id="27cf7-305">Для авторизации API, который отвечает за поток выполнения от имени другого субъекта, в случае надстроек Office используйте только область `access_as_user`. Для других API в службе должны быть предусмотрены отдельные требования, касающиеся областей.</span><span class="sxs-lookup"><span data-stu-id="27cf7-305">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements.</span></span> <span data-ttu-id="27cf7-306">Это ограничивает доступ, предоставляемый с использованием маркеров, которые получает Office.</span><span class="sxs-lookup"><span data-stu-id="27cf7-306">This limits what can be accessed with the tokens that Office acquires.</span></span>

5. <span data-ttu-id="27cf7-p160">Замените `TODO8` приведенным ниже кодом. Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p160">Replace `TODO8` with the following code. Note the following about this code:</span></span>

    * <span data-ttu-id="27cf7-309">Данные вызова `acquireTokenOnBehalfOf` не включают параметр ресурса, так как мы создали объект `AuthModule` (`auth`) с использованием конечной точки AAD версии 2.0, которая не поддерживает свойство ресурса.</span><span class="sxs-lookup"><span data-stu-id="27cf7-309">The call to `acquireTokenOnBehalfOf` does not include a resource parameter because we constructed the `AuthModule` object (`auth`) with the AAD V2.0 endpoint which does not support a resource property.</span></span>
    * <span data-ttu-id="27cf7-310">Второй параметр вызова задает разрешения, необходимые надстройке, чтобы получить список файлов и папок пользователя из OneDrive.</span><span class="sxs-lookup"><span data-stu-id="27cf7-310">The second parameter of the call specifies the permissions the add-in will need to get a list of the user's files and folders on OneDrive.</span></span> <span data-ttu-id="27cf7-311">(Разрешение `profile` не запрашивается, так как оно требуется, когда ведущее приложение Office получает маркер доступа к надстройке, а не когда вы меняете этот токен на маркер доступа к Microsoft Graph.)</span><span class="sxs-lookup"><span data-stu-id="27cf7-311">(The `profile` permission is not requested because it is only needed when the Office host gets the access token to your add-in, not when you are trading in that token for an access token to Microsoft Graph.)</span></span>

    ```typescript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

6. <span data-ttu-id="27cf7-p162">Замените `TODO9` приведенной ниже строкой. Обратите внимание на указанные ниже особенности этого кода.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p162">Replace `TODO9` with the following line. Note the following about this code:</span></span>

    * <span data-ttu-id="27cf7-314">Класс MSGraphHelper определен в файле src\msgraph-helper.ts.</span><span class="sxs-lookup"><span data-stu-id="27cf7-314">The MSGraphHelper class is defined in src\msgraph-helper.ts.</span></span> 
    * <span data-ttu-id="27cf7-315">Чтобы сократить количество возвращаемых данных, мы указываем, что нас интересуют только первые 3 элемента и свойство name.</span><span class="sxs-lookup"><span data-stu-id="27cf7-315">We minimize the data that must be returned by specifying that we only want the name property and only the first 3 items.</span></span>

    ```typescript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

7. <span data-ttu-id="27cf7-316">Замените `TODO10` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-316">Replace `TODO10` with the following code.</span></span> <span data-ttu-id="27cf7-317">Обратите внимание на то, что этот код обрабатывает ошибки "401 не санкционировано" Microsoft Graph, которые указывают на недействительный маркер или маркер с истекшим сроком действия.</span><span class="sxs-lookup"><span data-stu-id="27cf7-317">Note that this code handles '401 Unauthorized" errors from Microsoft Graph which would indicate an expired or invalid token.</span></span> <span data-ttu-id="27cf7-318">Вероятность такого события крайне мала, так как его должна предотвращать логика сохранения маркеров.</span><span class="sxs-lookup"><span data-stu-id="27cf7-318">It is very unlikely that this would ever happen since the token persisting logic should prevent it.</span></span> <span data-ttu-id="27cf7-319">(См. раздел **Создание метода для доступа к ресурсу с помощью потока "от имени"** выше.) Если это произойдет, код передаст клиенту ошибку с именем "Ошибка Microsoft Graph".</span><span class="sxs-lookup"><span data-stu-id="27cf7-319">(See the section **Create a method to get access to the resource using the "on behalf of" flow** above.) If it does happen, this code will relay the error to the client with "Microsoft Graph error" in the error name.</span></span> <span data-ttu-id="27cf7-320">(См. метод `handleClientSideErrors`, созданный вами в файле program.js на одном из более ранних этапов.) Код, который вы добавите в файл ODataHelper.js на одном из более поздних этапов, поможет обработать ошибки Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="27cf7-320">(See the `handleClientSideErrors` method that you created in the program.js file in an earlier step.) Code that you add to the ODataHelper.js file in a later step helps process errors from Microsoft Graph.</span></span>

    ```typescript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. <span data-ttu-id="27cf7-p164">Замените `TODO11` приведенным ниже кодом. Обратите внимание на то, что Microsoft Graph возвращает некоторые метаданные OData и свойство **eTag** для каждого элемента, даже если запрашивается только свойство `name`. Код отправляет клиенту только имена элементов.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p164">Replace `TODO11` with the following code. Note that Microsoft Graph returns some OData metadata and an **eTag** property for every item, even if `name` is the only property requested. The code sends only the item names to the client.</span></span>

    ```typescript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. <span data-ttu-id="27cf7-324">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="27cf7-324">Save and close the file.</span></span>

### <a name="add-response-handling-to-the-odatahelper"></a><span data-ttu-id="27cf7-325">Добавление обработки откликов в ODataHelper</span><span class="sxs-lookup"><span data-stu-id="27cf7-325">Add response handling to the ODataHelper</span></span>

1. <span data-ttu-id="27cf7-326">Откройте файл src\odata-helper.ts.</span><span class="sxs-lookup"><span data-stu-id="27cf7-326">Open the file src\odata-helper.ts.</span></span> <span data-ttu-id="27cf7-327">Файл почти завершен.</span><span class="sxs-lookup"><span data-stu-id="27cf7-327">The file is almost complete.</span></span> <span data-ttu-id="27cf7-328">Отсутствует текст обратного вызова обработчика для события завершения запроса.</span><span class="sxs-lookup"><span data-stu-id="27cf7-328">What's missing is the body of the callback to the handler for the request "end" event.</span></span> <span data-ttu-id="27cf7-329">Замените `TODO` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-329">Replace the `TODO` with the following code.</span></span> <span data-ttu-id="27cf7-330">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-330">About this code note:</span></span>

    * <span data-ttu-id="27cf7-331">Отклик от конечной точки OData может быть сообщением об ошибке, например 401, если конечная точка запрашивает маркер доступа, а он недействителен или срок его действия истек.</span><span class="sxs-lookup"><span data-stu-id="27cf7-331">The response from the OData endpoint might be an error, say a 401 if the endpoint requires an access token and it was invalid or expired.</span></span> <span data-ttu-id="27cf7-332">Но сообщение об ошибке по-прежнему является *сообщением*, а не ошибкой вызова `https.get`, поэтому строка `on('error', reject)` в конце `https.get` не запускается.</span><span class="sxs-lookup"><span data-stu-id="27cf7-332">But an error message is still a *message*, not an error in the call of `https.get`, so the `on('error', reject)` line at the end of `https.get` isn't triggered.</span></span> <span data-ttu-id="27cf7-333">Таким образом, код отличает сообщения об успешном выполнении (200) от сообщений об ошибках и отправляет объект JSON вызывающей стороне с запрошенными данными OData или информацией об ошибке.</span><span class="sxs-lookup"><span data-stu-id="27cf7-333">So, the code distinguishes success (200) messages from error messages and sends a JSON object to the caller with either the requested OData or error information.</span></span>

    ```typescript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1.  <span data-ttu-id="27cf7-p167">Замените `TODO1` приведенным ниже кодом. Обратите внимание: код предполагает, что данные возвращаются в формате JSON.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p167">Replace `TODO1` with the following code. Note that the code assumes the data is returned as JSON.</span></span>

    ```typescript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1.  <span data-ttu-id="27cf7-p168">Замените `TODO2` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="27cf7-p168">Replace `TODO2` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="27cf7-338">Отклик с сообщением об ошибке от источника OData будет иметь аргументы statusCode и statusMessage. При этом первый из них будет присутствовать всегда, а второй — обычно.</span><span class="sxs-lookup"><span data-stu-id="27cf7-338">An error response from an OData source will always have a statusCode and usually a statusMessage.</span></span> <span data-ttu-id="27cf7-339">Некоторые источники OData также добавляют в текст свойство ошибки с дополнительными сведениями, например внутренними данными или конкретизирующими сообщением и кодом.</span><span class="sxs-lookup"><span data-stu-id="27cf7-339">Some OData sources also add an error property to the body with further information, such as an inner, or more specific, code and message.</span></span>
    * <span data-ttu-id="27cf7-340">Объект Promise разрешен, не отклонен.</span><span class="sxs-lookup"><span data-stu-id="27cf7-340">The Promise object is resolved, not rejected.</span></span> <span data-ttu-id="27cf7-341">`https.get` выполняется, если веб-служба вызывает конечную точку OData "сервер-сервер".</span><span class="sxs-lookup"><span data-stu-id="27cf7-341">The `https.get` runs when a web service calls an OData endpoint server-to-server.</span></span> <span data-ttu-id="27cf7-342">Но этот вызов поступает в контексте вызова клиентом веб-API в веб-службе.</span><span class="sxs-lookup"><span data-stu-id="27cf7-342">But that call comes in the context of a call from a client to a web API in the web service.</span></span> <span data-ttu-id="27cf7-343">Если этот "внутренний" запрос отклонен, "внешний" запрос, отправляемый клиентом веб-службе, не выполняется.</span><span class="sxs-lookup"><span data-stu-id="27cf7-343">The "outer" request from the client to the web service never completes if this "inner" request is rejected.</span></span> <span data-ttu-id="27cf7-344">Кроме того, необходимо разрешить запрос с пользовательским объектом `Error`, если стороне, вызывающей `http.get`, необходимо передать клиенту сообщения об ошибках от конечной точки OData.</span><span class="sxs-lookup"><span data-stu-id="27cf7-344">Also, resolving the request with the custom `Error` object is required if the caller of `http.get` needs to relay errors from the OData endpoint to the client.</span></span>

    ```typescript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;
    
    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. <span data-ttu-id="27cf7-345">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="27cf7-345">Save and close the file.</span></span>

## <a name="deploy-the-add-in"></a><span data-ttu-id="27cf7-346">Развертывание надстройки</span><span class="sxs-lookup"><span data-stu-id="27cf7-346">Deploy the add-in</span></span>

<span data-ttu-id="27cf7-347">Теперь необходимо сообщить Office, где находится надстройка.</span><span class="sxs-lookup"><span data-stu-id="27cf7-347">Now you need to let Office know where to find the add-in.</span></span>

1. <span data-ttu-id="27cf7-348">Создайте сетевую папку или [предоставьте общий доступ к папке через сеть](https://docs.microsoft.com/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).</span><span class="sxs-lookup"><span data-stu-id="27cf7-348">Create a network share, or [share a folder to the network](https://docs.microsoft.com/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).</span></span>

2. <span data-ttu-id="27cf7-349">Поместите копию файла манифеста Office-Add-in-NodeJS-SSO.xml из корневой папки проекта в общую папку.</span><span class="sxs-lookup"><span data-stu-id="27cf7-349">Place a copy of the Office-Add-in-NodeJS-SSO.xml manifest file, from the root of the project, into the shared folder.</span></span>

3. <span data-ttu-id="27cf7-350">Запустите PowerPoint и откройте документ.</span><span class="sxs-lookup"><span data-stu-id="27cf7-350">Launch PowerPoint and open a document.</span></span>

4. <span data-ttu-id="27cf7-351">Перейдите на вкладку **Файл**, а затем выберите **Параметры**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-351">Choose the **File** tab, and then choose **Options**.</span></span>

5. <span data-ttu-id="27cf7-352">Выберите **Центр управления безопасностью**, а затем нажмите кнопку **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-352">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>

6. <span data-ttu-id="27cf7-353">Выберите пункт **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-353">Choose **Trusted Add-ins Catalogs**.</span></span>

7. <span data-ttu-id="27cf7-354">В поле **URL-адрес каталога** введите сетевой путь к общей папке с файлом Office-Add-in-NodeJS-SSO.xml и нажмите **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-354">In the **Catalog Url** field, enter the network path to the folder share that contains Office-Add-in-NodeJS-SSO.xml, and then choose **Add Catalog**.</span></span>

8. <span data-ttu-id="27cf7-355">Установите флажок **Показать в меню** и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-355">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

9. <span data-ttu-id="27cf7-p171">Появится сообщение о том, что параметры будут применены при следующем запуске Microsoft Office. Закройте PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p171">A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close PowerPoint.</span></span>

## <a name="build-and-run-the-project"></a><span data-ttu-id="27cf7-358">Сборка и запуск проекта</span><span class="sxs-lookup"><span data-stu-id="27cf7-358">Build and run the project</span></span>

<span data-ttu-id="27cf7-p172">Выполнить сборку проекта и запустить его можно двумя способами в зависимости от того, используете ли вы Visual Studio Code. В обоих случаях при каждом изменении кода автоматически выполняется повторная сборка, после чего проект запускается.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p172">There are two ways to build and run the project depending on whether you are using Visual Studio Code. For both ways, the project builds and automatically rebuilds and reruns when you make changes to the code.</span></span>

1. <span data-ttu-id="27cf7-361">Если вы не используете Visual Studio Code:</span><span class="sxs-lookup"><span data-stu-id="27cf7-361">If you are not using Visual Studio Code:</span></span> 
 1. <span data-ttu-id="27cf7-362">Откройте терминал node и перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="27cf7-362">Open a node terminal and navigate to the root folder of the project.</span></span>
 2. <span data-ttu-id="27cf7-363">Введите в терминале команду **npm run build**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-363">In the terminal, enter **npm run build**.</span></span> 
 3. <span data-ttu-id="27cf7-364">Откройте второй терминал node и перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="27cf7-364">Open a second node terminal and navigate to the root folder of the project.</span></span>
 4. <span data-ttu-id="27cf7-365">Введите в терминале команду **npm run start**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-365">In the terminal, enter **npm run start**.</span></span>

2. <span data-ttu-id="27cf7-366">Если используется VS Code:</span><span class="sxs-lookup"><span data-stu-id="27cf7-366">If you are using VS Code:</span></span>
 1. <span data-ttu-id="27cf7-367">Откройте проект в VS Code.</span><span class="sxs-lookup"><span data-stu-id="27cf7-367">Open the project in VS Code.</span></span>
 2. <span data-ttu-id="27cf7-368">Нажмите клавиши CTRL+SHIFT+B, чтобы выполнить сборку проекта.</span><span class="sxs-lookup"><span data-stu-id="27cf7-368">Press CTRL-SHIFT-B to build the project.</span></span>
 3. <span data-ttu-id="27cf7-369">Нажмите клавишу F5, чтобы запустить проект в сеансе отладки.</span><span class="sxs-lookup"><span data-stu-id="27cf7-369">Press F5 to run the project in a debugging session.</span></span>


## <a name="add-the-add-in-to-an-office-document"></a><span data-ttu-id="27cf7-370">Добавление надстройки в документ Office</span><span class="sxs-lookup"><span data-stu-id="27cf7-370">Add the add-in to an Office document</span></span>

1. <span data-ttu-id="27cf7-371">Перезапустите PowerPoint и откройте или создайте презентацию.</span><span class="sxs-lookup"><span data-stu-id="27cf7-371">Restart PowerPoint and open or create a presentation.</span></span>

1. <span data-ttu-id="27cf7-372">Если вкладка **Разработчик** не отображается на ленте, включите ее с помощью следующих действий:</span><span class="sxs-lookup"><span data-stu-id="27cf7-372">If the **Developer** tab is not visible on the ribbon, enable it with the following steps:</span></span>
 1. <span data-ttu-id="27cf7-373">Перейдите в раздел **Файл** | **Параметры** | **Настройка ленты**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-373">Navigate to **File** | **Options** | **Customize Ribbon**.</span></span>
 2. <span data-ttu-id="27cf7-374">Установите флажок, чтобы включить **разработчик** в дереве имен элементов управления в правой части страницы **Настройка ленты**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-374">Click the check box to enable **Developer** in the tree of control names on the right of the **Customize Ribbon** page.</span></span>
 3. <span data-ttu-id="27cf7-375">Нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-375">Press **OK**.</span></span>

2. <span data-ttu-id="27cf7-376">На вкладке **Разработчик** в PowerPoint выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-376">On the **Developer** tab in PowerPoint, choose **My Add-ins**.</span></span>

3. <span data-ttu-id="27cf7-377">Откройте вкладку **Общая папка**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-377">Select the **SHARED FOLDER** tab.</span></span>

4. <span data-ttu-id="27cf7-378">Выберите **SSO NodeJS Sample** и нажмите **ОК**.</span><span class="sxs-lookup"><span data-stu-id="27cf7-378">Choose **SSO NodeJS Sample**, and then select **OK**.</span></span>

5. <span data-ttu-id="27cf7-379">На ленте **Главная** появится новая группа **SSO NodeJS** с кнопкой **Show Add-in** (Показать надстройку) и значком.</span><span class="sxs-lookup"><span data-stu-id="27cf7-379">On the **Home** ribbon is a new group called **SSO NodeJS** with a button labeled **Show Add-in** and an icon.</span></span> 

## <a name="test-the-add-in"></a><span data-ttu-id="27cf7-380">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="27cf7-380">Test the add-in</span></span>

1. <span data-ttu-id="27cf7-381">Убедитесь в наличии нескольких файлов в OneDrive, чтобы можно было проверить результаты.</span><span class="sxs-lookup"><span data-stu-id="27cf7-381">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

2. <span data-ttu-id="27cf7-382">Нажмите кнопку **Show Add-in** (Показать надстройку), чтобы открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="27cf7-382">Click **Show Add-in** button to open the add-in.</span></span>

2. <span data-ttu-id="27cf7-p173">Откроется страница приветствия. Нажмите кнопку **Get my files from OneDrive** (Получить мои файлы из OneDrive).</span><span class="sxs-lookup"><span data-stu-id="27cf7-p173">The add-in opens with a Welcome page. Click the **Get My Files from OneDrive** button.</span></span>

2. <span data-ttu-id="27cf7-p174">Если вы вошли в Office, под кнопкой появится список ваших файлов и папок из OneDrive. В первый раз это может занять более 15 секунд.</span><span class="sxs-lookup"><span data-stu-id="27cf7-p174">If you are are signed into Office, a list of your files and folders on OneDrive will appear below the button. This may take more than 15 seconds the first time.</span></span>

3. <span data-ttu-id="27cf7-387">Если вы не вошли в Office, откроется всплывающее окно с предложением войти.</span><span class="sxs-lookup"><span data-stu-id="27cf7-387">If you are not signed into Office, a popup will open and prompt you to sign in.</span></span> <span data-ttu-id="27cf7-388">Список файлов и папок появится через несколько секунд после входа.</span><span class="sxs-lookup"><span data-stu-id="27cf7-388">After you have completed the sign-in, the list of your files and folders will appear after a few seconds.</span></span> <span data-ttu-id="27cf7-389">*Повторно нажимать кнопку не следует.*</span><span class="sxs-lookup"><span data-stu-id="27cf7-389">*You should not press the button a second time.*</span></span>

> [!NOTE]
> <span data-ttu-id="27cf7-390">Если вы ранее выполняли вход в Office с использованием другого идентификатора и все еще не закрыли некоторые из открытых тогда приложений Office, Office может не сменить идентификатор (даже если кажется, что это сделано для PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="27cf7-390">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="27cf7-391">Если это произойдет, возможен сбой при вызове Microsoft Graph или возврат данных для другого идентификатора.</span><span class="sxs-lookup"><span data-stu-id="27cf7-391">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="27cf7-392">Чтобы избежать этого, *закройте все приложения Office*, прежде чем нажимать кнопку **Get My Files from OneDrive** (Получить мои файлы из OneDrive).</span><span class="sxs-lookup"><span data-stu-id="27cf7-392">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>
