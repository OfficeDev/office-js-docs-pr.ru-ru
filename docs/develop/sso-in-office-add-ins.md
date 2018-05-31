---
title: Включение единого входа для надстроек Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 45bd63150ffa8e46bf9c0fa54711ac907b8490ce
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437516"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="2bc6f-102">Включение единого входа для надстроек Office (тестовый режим)</span><span class="sxs-lookup"><span data-stu-id="2bc6f-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="2bc6f-103">Пользователи входят в Office (в Интернете, на мобильных устройствах и настольных компьютерах), используя личную учетную запись Майкрософт либо рабочую или учебную учетную запись (Office 365).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="2bc6f-104">Воспользуйтесь удобной функцией единого входа для однократной авторизации пользователя в своей надстройке без необходимости повторного входа.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>


![Изображение, иллюстрирующее процесс входа в надстройку](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> <span data-ttu-id="2bc6f-106">В настоящее время API единого входа поддерживается для Word, Excel, Outlook и PowerPoint в тестовом режиме.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-106">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="2bc6f-107">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API идентификации](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-107">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).</span></span>
> <span data-ttu-id="2bc6f-108">Если вы работаете с надстройкой Outlook, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="2bc6f-109">Инструкции см. в [этой статье](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-109">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="2bc6f-110">Пользователям будет удобнее запускать надстройку, так как не придется каждый раз выполнять вход.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-110">For users, this makes running your add-in a smooth experience that involves at signing in only once.</span></span> <span data-ttu-id="2bc6f-111">Для разработчиков это означает, что в надстройке не будут храниться таблицы пользователей с зашифрованными паролями.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-111">For developers, this means that your add-in does not have to maintain it's own user tables with encrypted passwords.</span></span>

### <a name="how-it-works-at-runtime"></a><span data-ttu-id="2bc6f-112">Принцип работы во время выполнения</span><span class="sxs-lookup"><span data-stu-id="2bc6f-112">How it works at runtime</span></span>

<span data-ttu-id="2bc6f-113">На приведенной ниже схеме показано, как работает единый вход.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-113">The following diagram shows how the SSO process works.</span></span>

![Схема единого входа](../images/sso-overview-diagram.png)

1. <span data-ttu-id="2bc6f-115">Код JavaScript надстройки вызывает новый API Office.js — `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-115">In the add-in, JavaScript calls a new Office.js API `getAccessTokenAsync`.</span></span> <span data-ttu-id="2bc6f-116">Он указывает ведущему приложению Office, что необходимо получить маркер доступа к надстройке.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-116">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="2bc6f-117">См. раздел [Пример маркера доступа](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-117">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="2bc6f-118">Если вход в Office не выполнен, в ведущем приложении открывается всплывающее окно, в котором пользователю предлагается войти.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-118">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="2bc6f-119">Если пользователь запускает надстройку в первый раз, ему предлагается дать согласие.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-119">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="2bc6f-120">Ведущее приложение Office запрашивает **маркер надстройки** у конечной точки Azure AD версии 2.0 для текущего пользователя.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-120">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="2bc6f-121">Azure AD отправляет маркер надстройки ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-121">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="2bc6f-122">Ведущее приложение Office отправляет **маркер** надстройке в составе объекта результата, возвращенного при вызове метода `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-122">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="2bc6f-123">JavaScript в надстройке может проанализировать маркер и извлечь необходимую информацию, например, адрес электронной почты пользователя.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-123">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="2bc6f-124">Кроме того, надстройка может отправить HTTP-запрос на сервер для получения дополнительных сведений о пользователе, например, предпочтений пользователя.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-124">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="2bc6f-125">Можно также отправить маркер доступа на сервер для анализа и проверки.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-125">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="2bc6f-126">Разработка надстройки с единым входом</span><span class="sxs-lookup"><span data-stu-id="2bc6f-126">Develop an SSO add-in</span></span>

<span data-ttu-id="2bc6f-127">В этом разделе описаны задачи, необходимые для создания надстройки Office с единым входом.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-127">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="2bc6f-128">Эти задачи описываются независимо от языка и платформы.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-128">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="2bc6f-129">Подробные пошаговые инструкции см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="2bc6f-129">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="2bc6f-130">Создание надстройки Office на платформе Node.js с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="2bc6f-130">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="2bc6f-131">Создание надстройки Office на платформе ASP.NET с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="2bc6f-131">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="2bc6f-132">Создание приложения-службы</span><span class="sxs-lookup"><span data-stu-id="2bc6f-132">Create the service application</span></span>

<span data-ttu-id="2bc6f-133">Зарегистрируйте надстройку на портале регистрации для конечной точки Azure версии 2.0: https://apps.dev.microsoft.com. Этот процесс занимает 5–10 минут и включает следующие задачи.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-133">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="2bc6f-134">Получение идентификатора и секрета клиента для надстройки.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-134">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="2bc6f-135">Укажите необходимые для надстройки разрешения для конечной точки AAD версии</span><span class="sxs-lookup"><span data-stu-id="2bc6f-135">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="2bc6f-136">2.0 (и дополнительно для Microsoft Graph).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-136">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="2bc6f-137">Разрешение "профиля" требуется всегда.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-137">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="2bc6f-138">Предоставление надстройке доверия ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-138">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="2bc6f-139">Предварительная авторизация ведущего приложения Office для надстройки с помощью заданного по умолчанию разрешения *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-139">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="2bc6f-140">Для получения дополнительной информации об этом процессе см. статью [Регистрация надстройки Office, использующей единый вход с конечной точкой Azure AD версии 2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-140">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="2bc6f-141">Конфигурация надстройки</span><span class="sxs-lookup"><span data-stu-id="2bc6f-141">Configure the add-in</span></span>

<span data-ttu-id="2bc6f-142">Добавьте новую разметку в манифест надстройки:</span><span class="sxs-lookup"><span data-stu-id="2bc6f-142">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="2bc6f-143">**WebApplicationInfo** — родительский элемент для указанных ниже элементов.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-143">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="2bc6f-144">**Id** — идентификатор клиента надстройки. Это идентификатор приложения, который вы получаете в рамках регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-144">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="2bc6f-145">См. статью [Регистрация надстройки Office, использующей единый вход с конечной точкой Azure AD версии 2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-145">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="2bc6f-146">**Resource** — URL-адрес надстройки.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-146">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="2bc6f-147">**Scopes** — родительский элемент одного или нескольких элементов **Scope**.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-147">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="2bc6f-148">**Scope** — указывает разрешение, необходимое надстройке для работы с AAD.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-148">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="2bc6f-149">Разрешение `profile` требуется всегда, это может быть единственным необходимым разрешением, если надстройка не получает доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-149">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="2bc6f-150">Если надстройка получает этот доступ, потребуются элементы **Scope** для необходимых разрешений Microsoft Graph; например, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-150">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="2bc6f-151">Для библиотек, которые используются в коде для доступа к Microsoft Graph, могут потребоваться дополнительные разрешения.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-151">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="2bc6f-152">Например, для библиотеки проверки подлинности Майкрософт (MSAL) для .NET требуется разрешение `offline_access`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-152">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="2bc6f-153">Для получения дополнительной информации см. статью [Авторизованный доступ в Microsoft Graph из вашей надстройки Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-153">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="2bc6f-p110">Для всех ведущих приложений, кроме Outlook, добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Для Outlook добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-p110">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="2bc6f-156">Ниже приведен пример части кода.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-156">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="2bc6f-157">Добавление кода для клиента</span><span class="sxs-lookup"><span data-stu-id="2bc6f-157">Add client-side code</span></span>

<span data-ttu-id="2bc6f-158">Добавьте в надстройку код JavaScript для:</span><span class="sxs-lookup"><span data-stu-id="2bc6f-158">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="2bc6f-159">Вызовите [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-159">Call [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).</span></span>
* <span data-ttu-id="2bc6f-160">Проанализируйте маркер доступа или передайте его в серверный код надстройки.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-160">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="2bc6f-161">Далее представлен простой пример вызова `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-161">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!Note]
> <span data-ttu-id="2bc6f-162">В данном примере представлен только один тип ошибки.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-162">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="2bc6f-163">Примеры более сложной обработки ошибок: [Home.js в Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) и [program.js в Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-163">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="2bc6f-164">См. статью [Устранение ошибок единого входа](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-164">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

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

<span data-ttu-id="2bc6f-165">Далее представлен пример передачи маркера надстройки на сервер.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-165">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="2bc6f-166">При отправке запроса обратно на сервер маркер указывается в качестве заголовка `Authorization`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-166">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="2bc6f-167">Данный пример предусматривает отправку данных JSON, поэтому используется метод `POST`, однако `GET` достаточно для отправки маркера доступа, если не выполняется запись в сервер.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-167">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="2bc6f-168">Когда вызывать метод</span><span class="sxs-lookup"><span data-stu-id="2bc6f-168">When to call the method</span></span>

<span data-ttu-id="2bc6f-169">Если надстройка не может работать без входа в Office, необходимо вызвать `getAccessTokenAsync` *при запуске надстройки*.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-169">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="2bc6f-170">Если надстройка может работать без входа, метод `getAccessTokenAsync` *вызывается, когда требуется вход*.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-170">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="2bc6f-171">Производительность не снижается при повторяющихся вызовах `getAccessTokenAsync`, так как Office кэширует маркер доступа и использует его, пока не истечет срок его действия, не вызывая конечную точку AAD</span><span class="sxs-lookup"><span data-stu-id="2bc6f-171">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="2bc6f-172">версии 2.0 при каждом вызове `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-172">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="2bc6f-173">Поэтому вызовы `getAccessTokenAsync` можно добавлять во все функции и обработчики, которые инициируют действие, где нужен маркер.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-173">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="2bc6f-174">Добавление серверного кода</span><span class="sxs-lookup"><span data-stu-id="2bc6f-174">Add server-side code</span></span>

<span data-ttu-id="2bc6f-175">В большинстве случаев практически нет смысла получать маркер доступа, если надстройка не передает его на сторону сервера и не использует его там.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-175">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="2bc6f-176">Далее указаны некоторые серверные задачи, которые может выполнять надстройка.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-176">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="2bc6f-177">Создайте один или несколько методов веб-API, использующих информацию о пользователе, которая извлекается из маркера; например, метод поиска предпочтений пользователя в базе данных на сервере.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-177">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="2bc6f-178">(См. статью **Использование маркера единого входа в качестве удостоверения** далее.) В зависимости от языка и платформы могут быть доступны библиотеки, который упростят создание нужного кода.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-178">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="2bc6f-179">Получите данные Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-179">Get Microsoft Graph data.</span></span> <span data-ttu-id="2bc6f-180">Серверный код должен:</span><span class="sxs-lookup"><span data-stu-id="2bc6f-180">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="2bc6f-181">Проверка маркера доступа (см. статью **Проверка маркера доступа** далее).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-181">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="2bc6f-182">Запускать поток "от имени" путем вызова конечной точки Azure AD версии 2.0, включающего маркер доступа, некоторые метаданные пользователя и учетные данные надстройки (идентификатор и секрет).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-182">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="2bc6f-183">В этом контексте маркер доступа называется маркером начальной загрузки.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-183">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="2bc6f-184">Выполните кэширование нового маркера доступа от потока "от имени".</span><span class="sxs-lookup"><span data-stu-id="2bc6f-184">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="2bc6f-185">Получите данные от Microsoft Graph с помощью нового маркера.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-185">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="2bc6f-186">Для получения дополнительной информации о получении санкционированного доступа к данным пользователя Microsoft Graph см. статью [Авторизованный доступ в Microsoft Graph из вашей надстройки Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-186">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="2bc6f-187">Проверка маркера доступа</span><span class="sxs-lookup"><span data-stu-id="2bc6f-187">Validate the token</span></span>

<span data-ttu-id="2bc6f-188">Когда веб-API получит маркер доступа, этот токен необходимо проверить перед использованием.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-188">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="2bc6f-189">Это маркер JSON Web Token (JWT), то есть его проверка выполняется так же, как и в большинстве стандартных потоков OAuth.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-189">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="2bc6f-190">Доступно множество библиотек, которые могут выполнять проверку JWT, основные их действия:</span><span class="sxs-lookup"><span data-stu-id="2bc6f-190">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="2bc6f-191">проверяют правильность формата маркера;</span><span class="sxs-lookup"><span data-stu-id="2bc6f-191">Checking that the token is well-formed</span></span>
- <span data-ttu-id="2bc6f-192">проверяют, выдан ли маркер нужным центром сертификации;</span><span class="sxs-lookup"><span data-stu-id="2bc6f-192">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="2bc6f-193">проверяют, предназначен ли маркер для веб-API.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-193">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="2bc6f-194">При проверке маркера учитывайте приведенные ниже рекомендации.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-194">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="2bc6f-195">Действительные маркеры единого входа выдает центр сертификации Azure, `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-195">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="2bc6f-196">Утверждение `iss` в маркере должно начинаться с этого значения.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-196">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="2bc6f-197">Параметру `aud` маркера будет присвоено значение идентификатора приложения с портала регистрации.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-197">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="2bc6f-198">Для параметра `scp` маркера будет задано значение `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-198">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="2bc6f-199">Использование маркера единого входа в качестве удостоверения</span><span class="sxs-lookup"><span data-stu-id="2bc6f-199">Using the SSO token as an identity</span></span>

<span data-ttu-id="2bc6f-200">Если приложению необходимо проверить удостоверение пользователя, то маркер единого входа содержит сведения, с помощью которых можно определить его.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-200">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="2bc6f-201">Ниже перечислены утверждения из маркера, связанные с удостоверениями.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-201">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="2bc6f-202">`name` — Отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-202">`name` - The user's display name.</span></span>
- <span data-ttu-id="2bc6f-203">`preferred_username` — Адрес электронной почты пользователя.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-203">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="2bc6f-204">`oid` — GUID, предоставляющий ИД пользователя в Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-204">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="2bc6f-205">`tid` — GUID, предоставляющий ИД организации пользователя в Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-205">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="2bc6f-206">Значения `name` и `preferred_username` могут меняться, поэтому рекомендуется использовать значения `oid` и `tid`, чтобы связать удостоверение с внутренней службой авторизации.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-206">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="2bc6f-207">Например, если служба может форматировать эти значения вместе (в виде `{oid-value}@{tid-value}`), то их следует хранить в качестве значения в записи пользователя во внутренней базе данных пользователей.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-207">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="2bc6f-208">При последующих запросах удостоверение пользователя можно будет получать с помощью того же значения, а доступ к определенным ресурсам может предоставляться в соответствии с действующими механизмами управления доступом.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-208">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="2bc6f-209">Пример маркера доступа</span><span class="sxs-lookup"><span data-stu-id="2bc6f-209">Example access token</span></span>

<span data-ttu-id="2bc6f-210">Далее представлены типичные расшифрованные полезные данные маркера доступа.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-210">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="2bc6f-211">Для получения дополнительной информации о свойствах см. статью [Ссылка на маркеры Azure Active Directory v2.0](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-211">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-and-outlook-add-in"></a><span data-ttu-id="2bc6f-212">Использование единого входа в надстройке Outlook</span><span class="sxs-lookup"><span data-stu-id="2bc6f-212">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="2bc6f-213">Имеются небольшие, но важные различия между использованием функции единого входа в надстройке Outlook и использованием ее в надстройке Excel, PowerPoint или Word.</span><span class="sxs-lookup"><span data-stu-id="2bc6f-213">There are some small, but important differences in using SSO in and Outlook add-in from using it in as Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="2bc6f-214">Ознакомьтесь со статьями [Аутентификация пользователя с маркером единого входа в надстройке Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/authenticate-a-user-with-an-sso-token) и [Сценарий: реализация единого входа для службы в надстройке Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="2bc6f-214">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/en-us/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/en-us/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>