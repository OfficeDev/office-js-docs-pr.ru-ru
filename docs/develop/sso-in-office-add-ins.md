---
title: Включение единого входа для надстроек Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: dc9050d574e0a5e74ae8cae2c63817aa4f952eb9
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691197"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="813ad-102">Включение единого входа для надстроек Office (тестовый режим)</span><span class="sxs-lookup"><span data-stu-id="813ad-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="813ad-103">Пользователи входят в Office (в Интернете, на мобильных устройствах и настольных компьютерах), используя личную учетную запись Майкрософт либо рабочую или учебную учетную запись (Office 365).</span><span class="sxs-lookup"><span data-stu-id="813ad-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="813ad-104">Вы можете использовать эту возможность и использовать единый вход для авторизации пользователя в вашей надстройке, при этом пользователю не потребуется входить повторно.</span><span class="sxs-lookup"><span data-stu-id="813ad-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Изображение, иллюстрирующее процесс входа в надстройку](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="813ad-106">Предварительный просмотр состояния</span><span class="sxs-lookup"><span data-stu-id="813ad-106">Preview Status</span></span>

<span data-ttu-id="813ad-107">API единого входа в настоящее время поддерживается только для предварительного просмотра.</span><span class="sxs-lookup"><span data-stu-id="813ad-107">The Single Sign-on API is currently supported in preview only.</span></span> <span data-ttu-id="813ad-108">Также оно доступно для разработчиков для целей экспериментирования; но его нельзя использовать в рабочей надстройке.</span><span class="sxs-lookup"><span data-stu-id="813ad-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="813ad-109">Кроме того, надстройки, которые используют единый вход, не допускаются к использованию в [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="813ad-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="813ad-110">Для единого входа требуется Office 365 (версия Office, распространяемая по подписке).</span><span class="sxs-lookup"><span data-stu-id="813ad-110">SSO requires Office 365 (the subscription version of Office, also called “Click to Run”).</span></span> <span data-ttu-id="813ad-111">Следует использовать последнюю версию для текущего месяца и сборку из канала для участников программы предварительной оценки.</span><span class="sxs-lookup"><span data-stu-id="813ad-111">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="813ad-112">Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="813ad-112">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="813ad-113">Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="813ad-113">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="813ad-114">Обратите внимание на то, что когда сборка будет готова для выпуска на канале Semi-annual channel, поддержка функций предварительного просмотра, включая единый вход, отключается для этой сборки.</span><span class="sxs-lookup"><span data-stu-id="813ad-114">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

<span data-ttu-id="813ad-115">Не все приложения Office поддерживают предварительный просмотр единого входа.</span><span class="sxs-lookup"><span data-stu-id="813ad-115">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="813ad-116">Эта возможность доступна в Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="813ad-116">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="813ad-117">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API идентификации](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="813ad-117">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="813ad-118">Рекомендации и требования</span><span class="sxs-lookup"><span data-stu-id="813ad-118">Requirements and Best Practices</span></span>

<span data-ttu-id="813ad-119">Чтобы использовать единый вход, вам необходимо загрузить бета-версию библиотеки JavaScript Office из `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` на страницу подготовки HTML для надстройки.</span><span class="sxs-lookup"><span data-stu-id="813ad-119">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="813ad-120">Если вы работаете с надстройкой **Outlook**, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="813ad-120">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="813ad-121">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить в клиенте современную проверку подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="813ad-121">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="813ad-122">Вам *не* следует опираться на функцию единого в качестве единого способа проверки подлинности вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="813ad-122">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="813ad-123">Вам следует внедрить альтернативную систему проверки подлинности, на которую ваша надстройка может опираться при определенных ошибках.</span><span class="sxs-lookup"><span data-stu-id="813ad-123">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="813ad-124">Вы можете использовать систему таблиц пользователя и проверки подлинности, либо вы можете выделить одну из систем входа с использованием социальных сервисов.</span><span class="sxs-lookup"><span data-stu-id="813ad-124">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="813ad-125">Дополнительные сведения о том, как это сделать с помощью надстройки Office см. в статье [Авторизация внешних служб в надстройке Office](/office/dev/add-ins/develop/auth-external-add-ins).</span><span class="sxs-lookup"><span data-stu-id="813ad-125">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](/office/dev/add-ins/develop/auth-external-add-ins).</span></span> <span data-ttu-id="813ad-126">Для *Outlook*, существует рекомендуемая система возврата.</span><span class="sxs-lookup"><span data-stu-id="813ad-126">For *Outlook*, there is a recommended fall back system.</span></span> <span data-ttu-id="813ad-127">Дополнительные сведения см. в статье [Сценарий: реализация единого входа для службы в надстройке Outlook](/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="813ad-127">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="813ad-128">Принцип работы единого входа во время выполнения</span><span class="sxs-lookup"><span data-stu-id="813ad-128">How SSO works at runtime</span></span>

<span data-ttu-id="813ad-129">На приведенной ниже схеме показано, как работает единый вход.</span><span class="sxs-lookup"><span data-stu-id="813ad-129">The following diagram shows how the SSO process works.</span></span>

![Схема единого входа](../images/sso-overview-diagram.png)

1. <span data-ttu-id="813ad-131">В надстройке JavaScript вызывает новый API Office.js [getAccessTokenAsync](#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="813ad-131">In the add-in, JavaScript calls a new Office.js API [getAccessTokenAsync](#sso-api-reference).</span></span> <span data-ttu-id="813ad-132">Он указывает ведущему приложению Office на то, что необходимо получить маркер доступа для надстройки.</span><span class="sxs-lookup"><span data-stu-id="813ad-132">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="813ad-133">См. [Пример маркера доступа](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="813ad-133">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="813ad-134">Если вход в Office не выполнен, в ведущем приложении открывается всплывающее окно, в котором пользователю предлагается войти.</span><span class="sxs-lookup"><span data-stu-id="813ad-134">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="813ad-135">Если пользователь запускает надстройку в первый раз, ему предлагается дать согласие.</span><span class="sxs-lookup"><span data-stu-id="813ad-135">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="813ad-136">Ведущее приложение Office запрашивает **маркер надстройки** у конечной точки Azure AD версии 2.0 для текущего пользователя.</span><span class="sxs-lookup"><span data-stu-id="813ad-136">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="813ad-137">Azure AD отправляет маркер надстройки ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="813ad-137">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="813ad-138">Ведущее приложение Office отправляет **маркер** надстройке в составе объекта результата, возвращенного при вызове метода `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="813ad-138">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="813ad-139">JavaScript в надстройке может проанализировать маркер и извлечь данные, которые необходимы, например адрес электронной почты.</span><span class="sxs-lookup"><span data-stu-id="813ad-139">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="813ad-140">Кроме того, надстройка может отправить HTTP-запрос на серверную часть для получения дополнительных данных о пользователе, например, настройки пользователя.</span><span class="sxs-lookup"><span data-stu-id="813ad-140">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="813ad-141">Кроме того маркер доступа сам может быть отправлен на серверную часть для анализа и проверки.</span><span class="sxs-lookup"><span data-stu-id="813ad-141">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span>

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="813ad-142">Разработка надстройки с единым входом</span><span class="sxs-lookup"><span data-stu-id="813ad-142">Develop an SSO add-in</span></span>

<span data-ttu-id="813ad-p109">В этом разделе описаны задачи, необходимые для создания надстройки Office с единым входом. Эти задачи описываются независимо от языка и платформы. Подробные пошаговые инструкции см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="813ad-p109">This section describes the tasks involved in creating an Office Add-in that uses SSO. These tasks are described here in a language- and framework-agnostic way. For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="813ad-146">Создание надстройки Office на платформе Node.js с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="813ad-146">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="813ad-147">Создание надстройки Office на платформе ASP.NET с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="813ad-147">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="813ad-148">Создание приложения-службы</span><span class="sxs-lookup"><span data-stu-id="813ad-148">Create the service application</span></span>

<span data-ttu-id="813ad-p110">Зарегистрируйте надстройку на портале регистрации для конечной точки Azure версии 2.0: https://apps.dev.microsoft.com. Этот процесс занимает 5–10 минут и включает следующие задачи:</span><span class="sxs-lookup"><span data-stu-id="813ad-p110">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="813ad-151">Получение идентификатора и секрета клиента для надстройки;</span><span class="sxs-lookup"><span data-stu-id="813ad-151">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="813ad-152">Указание необходимых надстройке разрешений для конечной точки AAD в.</span><span class="sxs-lookup"><span data-stu-id="813ad-152">Specify the permissions that your add-in needs to AAD v.</span></span> <span data-ttu-id="813ad-153">2.0 (и Microsoft Graph в качестве опции).</span><span class="sxs-lookup"><span data-stu-id="813ad-153">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="813ad-154">Разрешение «профиль» необходимо во всех случаях.</span><span class="sxs-lookup"><span data-stu-id="813ad-154">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="813ad-155">Предоставление надстройке доверия ведущего приложения Office;</span><span class="sxs-lookup"><span data-stu-id="813ad-155">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="813ad-156">Предварительная авторизация ведущего приложения Office для надстройки с помощью заданного по умолчанию разрешения *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="813ad-156">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="813ad-157">Дополнительные сведения о данном процессе см. в статье [Регистрация надстройки Office, использующей единый вход с конечной точкой Microsoft Azure AD версии 2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="813ad-157">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="813ad-158">Конфигурация надстройки</span><span class="sxs-lookup"><span data-stu-id="813ad-158">Configure the add-in</span></span>

<span data-ttu-id="813ad-159">Добавьте новую разметку в манифест надстройки:</span><span class="sxs-lookup"><span data-stu-id="813ad-159">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="813ad-160">**WebApplicationInfo** — родительский элемент для указанных ниже элементов.</span><span class="sxs-lookup"><span data-stu-id="813ad-160">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="813ad-161">**Id** - Идентификатор клиента надстройки. Это идентификатор приложения, который вы получаете в процессе регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="813ad-161">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="813ad-162">См. [Регистрация надстройки Office, использующей единый вход с конечной точкой Microsoft Azure AD версии 2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="813ad-162">See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="813ad-163">**Resource** — URL-адрес надстройки;</span><span class="sxs-lookup"><span data-stu-id="813ad-163">**Resource** - The URL of the add-in.</span></span> <span data-ttu-id="813ad-164">Это тот же URI (включая протокол `api:`), который вы использовали при регистрации надстройки и в AAD.</span><span class="sxs-lookup"><span data-stu-id="813ad-164">This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD.</span></span> <span data-ttu-id="813ad-165">Доменная часть данного URI должна соответствовать домену, в том числе поддомену, используемом в URL-адресах в части`<Resources>` манифеста настройки.</span><span class="sxs-lookup"><span data-stu-id="813ad-165">The domain part of this URI should match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>
* <span data-ttu-id="813ad-166">**Scopes** — родительский элемент одного или нескольких элементов **Scope**;</span><span class="sxs-lookup"><span data-stu-id="813ad-166">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="813ad-167">**Scope** — указывает разрешение, необходимое надстройке для работы с AAD.</span><span class="sxs-lookup"><span data-stu-id="813ad-167">**Scope** - Specifies a permission that the add-in needs to AAD.</span></span> <span data-ttu-id="813ad-168">Разрешение `profile` требуется во всех случаях, и оно может быть единственным необходимым разрешением, если ваша надстройка не имеет доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="813ad-168">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="813ad-169">В противном случае вам также могут потребоваться элементы типа **Область** для необходимым разрешений Microsoft Graph; например, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="813ad-169">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="813ad-170">Библиотеки, которые вы используете в коде, чтобы получить доступ к Microsoft Graph, могут потребовать дополнительные разрешения.</span><span class="sxs-lookup"><span data-stu-id="813ad-170">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="813ad-171">Например, библиотека проверки подлинности Microsoft (MSAL) для .NET требует разрешения `offline_access`.</span><span class="sxs-lookup"><span data-stu-id="813ad-171">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="813ad-172">Дополнительные сведения см. в статье [Авторизация в Microsoft Graph для надстройки Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="813ad-172">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="813ad-p115">Для всех ведущих приложений, кроме Outlook, добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Для Outlook добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="813ad-p115">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="813ad-175">Ниже приведен пример разметки:</span><span class="sxs-lookup"><span data-stu-id="813ad-175">The following is an example of the markup:</span></span>

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### <a name="add-client-side-code"></a><span data-ttu-id="813ad-176">Добавьте код для клиента</span><span class="sxs-lookup"><span data-stu-id="813ad-176">Add client-side code</span></span>

<span data-ttu-id="813ad-177">Добавьте в надстройку код JavaScript для:</span><span class="sxs-lookup"><span data-stu-id="813ad-177">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="813ad-178">Вызова [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)</span><span class="sxs-lookup"><span data-stu-id="813ad-178">Call [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="813ad-179">Выполните анализ маркера доступа или передайте его в код надстройки на стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="813ad-179">Parse the access token or pass it to the add-in’s server-side code.</span></span>

<span data-ttu-id="813ad-180">Вот простой пример вызова для `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="813ad-180">Here's a simple example of a call to `getAccessTokenAsync`.</span></span>

> [!NOTE]
> <span data-ttu-id="813ad-181">В этом примере обрабатывается только один тип ошибки явным образом.</span><span class="sxs-lookup"><span data-stu-id="813ad-181">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="813ad-182">Примеры более сложной обработки ошибок см. в статье [Home.js в Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) и [program.js в Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span><span class="sxs-lookup"><span data-stu-id="813ad-182">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="813ad-183">См. также [Устранение ошибок единого входа](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="813ad-183">And see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).</span></span>
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

<span data-ttu-id="813ad-184">Вот простой пример передачи маркера надстройки стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="813ad-184">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="813ad-185">Маркер включается в виде заголовка `Authorization` при отправке запроса назад стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="813ad-185">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="813ad-186">В этом примере представлена отправка данных JSON, поэтому в нем используется метод `POST`, но `GET` достаточно для отправки маркера доступа, когда запись не осуществляется на сервере.</span><span class="sxs-lookup"><span data-stu-id="813ad-186">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### <a name="when-to-call-the-method"></a><span data-ttu-id="813ad-187">Когда вызывать метод</span><span class="sxs-lookup"><span data-stu-id="813ad-187">When to call the method</span></span>

<span data-ttu-id="813ad-188">Если надстройка не может работать, когда ни один пользователь не выполнил вход в Office, тогда необходимо вызывать метод `getAccessTokenAsync` *при запуске надстройки*.</span><span class="sxs-lookup"><span data-stu-id="813ad-188">If your add-in cannot be used when no user is logged into Office, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="813ad-189">Если ряд функций надстройки может работать без входа пользователя в систему, вы вызываете `getAccessTokenAsync` \*, когда пользователь выполняет действие, для которого необходимо выполнить вход \*.</span><span class="sxs-lookup"><span data-stu-id="813ad-189">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires a logged in user*.</span></span> <span data-ttu-id="813ad-190">Производительность не снижается при повторяющихся вызовах `getAccessTokenAsync`, так как Office кэширует маркер доступа и использует его, пока не истечет срок его действия, не вызывая конечную точку AAD в.</span><span class="sxs-lookup"><span data-stu-id="813ad-190">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD v.</span></span> <span data-ttu-id="813ad-191">2.0 при каждом вызове `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="813ad-191">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="813ad-192">Поэтому вызовы `getAccessTokenAsync` можно добавлять во все функции и обработчики, которые инициируют действие, где нужен маркер.</span><span class="sxs-lookup"><span data-stu-id="813ad-192">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="813ad-193">Добавление серверного кода</span><span class="sxs-lookup"><span data-stu-id="813ad-193">Add server-side code</span></span>

<span data-ttu-id="813ad-194">В большинстве случаев практически нет смысла получать маркер доступа, если надстройка не передает его на сторону сервера и не использует его там.</span><span class="sxs-lookup"><span data-stu-id="813ad-194">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="813ad-195">Некоторые задачи на стороне сервера, которые может выполнять ваша надстройка:</span><span class="sxs-lookup"><span data-stu-id="813ad-195">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="813ad-196">Создание одного или нескольких методов Web API, которые используют сведения о пользователе, извлеченные из маркера, например, метод, который выполняет поиск параметров пользователя в вашей базе данных.</span><span class="sxs-lookup"><span data-stu-id="813ad-196">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="813ad-197">(См. **Использование маркера единого входа в качестве удостоверения** ниже) В зависимости от языка и платформы, могут быть доступны библиотеки, который упростят создание нужного кода.</span><span class="sxs-lookup"><span data-stu-id="813ad-197">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="813ad-198">Получение данных Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="813ad-198">Get Microsoft Graph data.</span></span> <span data-ttu-id="813ad-199">Ваш серверный код должен выполнять следующее:</span><span class="sxs-lookup"><span data-stu-id="813ad-199">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="813ad-200">Проверка маркера доступа (см. статью **Проверка маркер доступа** ниже).</span><span class="sxs-lookup"><span data-stu-id="813ad-200">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="813ad-201">Запуск потока "от имени" с помощью вызова конечной точки Microsoft Azure AD версии 2.0, включающего маркер доступа, некоторые метаданные пользователя и учетные данные надстройки (идентификатор и секрет).</span><span class="sxs-lookup"><span data-stu-id="813ad-201">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="813ad-202">В этом контексте маркер доступа вызывается маркером начальной загрузки.</span><span class="sxs-lookup"><span data-stu-id="813ad-202">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="813ad-203">Кэширование нового маркера доступа, полученного из потока «от-имени».</span><span class="sxs-lookup"><span data-stu-id="813ad-203">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="813ad-204">Получение данных из Microsoft Graph с помощью нового маркера.</span><span class="sxs-lookup"><span data-stu-id="813ad-204">Get data from Microsoft Graph by using the new token.</span></span>

 <span data-ttu-id="813ad-205">Дополнительные сведения о получении авторизованного доступа к данным Microsoft Graph пользователя см. в статье [Авторизация в Microsoft Graph для надстройки Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="813ad-205">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="813ad-206">Проверка маркера доступа</span><span class="sxs-lookup"><span data-stu-id="813ad-206">Validate the access token</span></span>

<span data-ttu-id="813ad-207">Когда Web API получит маркер доступа, этот токен необходимо проверить перед использованием.</span><span class="sxs-lookup"><span data-stu-id="813ad-207">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="813ad-208">Это маркер JSON Web Token (JWT), то есть его проверка выполняется так же, как и в большинстве стандартных потоков OAuth.</span><span class="sxs-lookup"><span data-stu-id="813ad-208">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="813ad-209">Доступно множество библиотек, которые могут выполнять проверку JWT, основные их действия:</span><span class="sxs-lookup"><span data-stu-id="813ad-209">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="813ad-210">проверяют правильность формата маркера;</span><span class="sxs-lookup"><span data-stu-id="813ad-210">Checking that the token is well-formed</span></span>
- <span data-ttu-id="813ad-211">проверяют, выдан ли маркер нужным центром сертификации;</span><span class="sxs-lookup"><span data-stu-id="813ad-211">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="813ad-212">проверяют, предназначен ли маркер для веб-API.</span><span class="sxs-lookup"><span data-stu-id="813ad-212">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="813ad-213">При проверке маркера учитывайте приведенные ниже рекомендации.</span><span class="sxs-lookup"><span data-stu-id="813ad-213">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="813ad-214">Действительные маркеры единого входа выдает центр сертификации Azure, `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="813ad-214">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="813ad-215">Утверждение `iss` в маркере должно начинаться с этого значения.</span><span class="sxs-lookup"><span data-stu-id="813ad-215">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="813ad-216">Параметру `aud` маркера будет присвоено значение идентификатора приложения с портала регистрации.</span><span class="sxs-lookup"><span data-stu-id="813ad-216">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="813ad-217">Для параметра `scp` маркера будет задано значение `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="813ad-217">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="813ad-218">Использование маркера единого входа в качестве удостоверения</span><span class="sxs-lookup"><span data-stu-id="813ad-218">Using the SSO token as an identity</span></span>

<span data-ttu-id="813ad-219">Если приложению необходимо проверить удостоверение пользователя, то маркер единого входа содержит сведения, с помощью которых можно определить его.</span><span class="sxs-lookup"><span data-stu-id="813ad-219">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="813ad-220">Ниже перечислены утверждения из маркера, связанные с удостоверениями.</span><span class="sxs-lookup"><span data-stu-id="813ad-220">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="813ad-221">`name` — Отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="813ad-221">`name` - The user's display name.</span></span>
- <span data-ttu-id="813ad-222">`preferred_username` — Электронный адрес пользователя.</span><span class="sxs-lookup"><span data-stu-id="813ad-222">`preferred_username` - The user's email address.</span></span>
- <span data-ttu-id="813ad-223">`oid` — GUID, предоставляющий ИД пользователя в Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="813ad-223">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="813ad-224">`tid` — GUID, предоставляющий ИД организации пользователя в Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="813ad-224">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="813ad-225">Значения `name` и `preferred_username` могут меняться, мы рекомендуем использовать значения `oid` и `tid`, чтобы связать удостоверение с внутренней службой авторизации.</span><span class="sxs-lookup"><span data-stu-id="813ad-225">Since the `name` and `preferred_username` values could change, we recommend that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="813ad-226">Например, если служба может форматировать эти значения вместе (в виде `{oid-value}@{tid-value}`), то их следует хранить в качестве значения в записи пользователя во внутренней базе данных пользователей.</span><span class="sxs-lookup"><span data-stu-id="813ad-226">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="813ad-227">При последующих запросах удостоверение пользователя можно будет получать с помощью того же значения, а доступ к определенным ресурсам может предоставляться в соответствии с действующими механизмами управления доступом.</span><span class="sxs-lookup"><span data-stu-id="813ad-227">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="813ad-228">Пример маркера доступа</span><span class="sxs-lookup"><span data-stu-id="813ad-228">Example access token</span></span>

<span data-ttu-id="813ad-229">Ниже приведен типичная раскодированная нагрузка маркера доступа.</span><span class="sxs-lookup"><span data-stu-id="813ad-229">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="813ad-230">Сведения о свойствах см. в статье [Справочные материалы для маркеров в Azure Active Directory 2.0](/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="813ad-230">For information about the properties, see [Azure Active Directory v2.0 tokens reference](/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",
    iat: 1521143967,
    nbf: 1521143967,
    exp: 1521147867,
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",
    azpacr: "0",
    e_exp: 262800,
    name: "Mila Nikolova",
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",
    preferred_username: "milan@contoso.com",
    scp: "access_as_user",
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",
    uti: "MICAQyhrH02ov54bCtIDAA",
    ver: "2.0"
}
```

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="813ad-231">С использованием единого входа с надстройкой Outlook</span><span class="sxs-lookup"><span data-stu-id="813ad-231">Using SSO with an Outlook add-in</span></span>

<span data-ttu-id="813ad-232">Существует ряд небольшие, но важных различий при использовании единого входа в надстройке Outlook и его использования в надстройках Excel, PowerPoint или Word.</span><span class="sxs-lookup"><span data-stu-id="813ad-232">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="813ad-233">Обязательно ознакомьтесь с содержанием статей [Выполнение проверки подлинности пользователя с маркером единого входа в надстройке Outlook](/outlook/add-ins/authenticate-a-user-with-an-sso-token) и [Сценарий: Реализация единого входа для вашей службы в надстройке Outlook](/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="813ad-233">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="813ad-234">Справочные материалы по API единого входа</span><span class="sxs-lookup"><span data-stu-id="813ad-234">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="813ad-235">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="813ad-235">getAccessTokenAsync</span></span>

<span data-ttu-id="813ad-236">В пространстве имен Office Auth, `Office.context.auth`, имеется метод,`getAccessTokenAsync`, который позволяет узлу Office получать маркер доступа для надстройки веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="813ad-236">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="813ad-237">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="813ad-237">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="813ad-238">Метод вызывает конечную точку Azure Active Directory версии 2.0, чтобы получить маркер доступа к вашей надстройке в веб-приложении.</span><span class="sxs-lookup"><span data-stu-id="813ad-238">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="813ad-239">Это позволяет надстройкам идентифицировать пользователей.</span><span class="sxs-lookup"><span data-stu-id="813ad-239">This enables add-ins to identify users.</span></span> <span data-ttu-id="813ad-240">Код на стороне сервера может использовать этот маркер для доступа к Microsoft Graph, чтобы добавить веб-приложение надстройки с помощью [потока OAuth "от имени пользователя"](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="813ad-240">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="813ad-241">В Outlook эта API не поддерживается, если надстройка загружается в почтовый ящик Outlook.com или Gmail.</span><span class="sxs-lookup"><span data-stu-id="813ad-241">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

|<span data-ttu-id="813ad-242">Узлы</span><span class="sxs-lookup"><span data-stu-id="813ad-242">Hosts</span></span>|<span data-ttu-id="813ad-243">Excel, OneNote, Outlook, PowerPoint и Word.</span><span class="sxs-lookup"><span data-stu-id="813ad-243">Excel, OneNote, Outlook, PowerPoint, Word</span></span>|
|---|---|
|[<span data-ttu-id="813ad-244">Наборы требований</span><span class="sxs-lookup"><span data-stu-id="813ad-244">Requirement sets</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)|[<span data-ttu-id="813ad-245">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="813ad-245">IdentityAPI</span></span>](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)|

#### <a name="parameters"></a><span data-ttu-id="813ad-246">Параметры</span><span class="sxs-lookup"><span data-stu-id="813ad-246">Parameters</span></span>

<span data-ttu-id="813ad-247">`options` - Опционально.</span><span class="sxs-lookup"><span data-stu-id="813ad-247">`options` - Optional.</span></span> <span data-ttu-id="813ad-248">Принимает объект `AuthOptions` (см. ниже) для определения поведения при входе.</span><span class="sxs-lookup"><span data-stu-id="813ad-248">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="813ad-249">`callback` - Опционально.</span><span class="sxs-lookup"><span data-stu-id="813ad-249">`callback` - Optional.</span></span> <span data-ttu-id="813ad-250">Принимает метод обратного вызова, который может выполнить анализ маркера для идентификатора пользователя или использовать маркер в потоке «от имени ваших», чтобы получать доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="813ad-250">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="813ad-251">Если [AsyncResult](/javascript/api/office/office.asyncresult) `.status` был выполнен «успешно», тогда `AsyncResult.value` представляет собой необработанный маркер доступа AAD</span><span class="sxs-lookup"><span data-stu-id="813ad-251">If [AsyncResult](/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="813ad-252">версии 2.0.</span><span class="sxs-lookup"><span data-stu-id="813ad-252">2.0-formatted access token.</span></span>

<span data-ttu-id="813ad-253">Интерфейс `AuthOptions` предоставляет опции для взаимодействия с пользователем, когда Office получает маркер доступа для надстройки из AAD в.</span><span class="sxs-lookup"><span data-stu-id="813ad-253">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="813ad-254">2.0 с методом `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="813ad-254">2.0 with the `getAccessTokenAsync` method.</span></span>

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try"
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```
