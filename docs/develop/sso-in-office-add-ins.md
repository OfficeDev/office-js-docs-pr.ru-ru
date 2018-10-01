---
title: Включение единого входа для надстроек Office
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: 05b5088a61df3f77a09b60dbdc3129074d5f8530
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348172"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="99999-102">Включение единого входа для надстроек Office (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="99999-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="99999-103">Пользователи входят в Office (онлайн, на мобильной или настольной платформе), используя личную, рабочую или учебную учетную запись Майкрософт (Office 365).</span><span class="sxs-lookup"><span data-stu-id="99999-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="99999-104">Воспользуйтесь удобной функцией единого входа для однократной авторизации пользователя в своей надстройке без необходимости повторного входа.</span><span class="sxs-lookup"><span data-stu-id="99999-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Изображение, иллюстрирующее процесс входа в надстройку](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="99999-106">Статус предварительной версии</span><span class="sxs-lookup"><span data-stu-id="99999-106">Preview Status</span></span>

<span data-ttu-id="99999-107">API единого входа в настоящее время поддерживается только в предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="99999-107">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="99999-108">Он доступен разработчикам в экспериментальных целях. Его не следует применять в рабочих надстройках.</span><span class="sxs-lookup"><span data-stu-id="99999-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="99999-109">Кроме того, надстройки, в которых используется единый вход, не принимаются в [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="99999-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="99999-110">Предварительную версию службы единого входа поддерживают не все приложения Office.</span><span class="sxs-lookup"><span data-stu-id="99999-110">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="99999-111">Она доступна для Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="99999-111">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="99999-112">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API идентификации](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="99999-112">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="99999-113">Требования и рекомендации</span><span class="sxs-lookup"><span data-stu-id="99999-113">Requirements and Best Practices</span></span>

<span data-ttu-id="99999-114">Чтобы использовать единый вход, необходимо подключить бета-версию библиотеки JavaScript для Office из `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` на начальной HTML-странице надстройки.</span><span class="sxs-lookup"><span data-stu-id="99999-114">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="99999-115">Если вы работаете с надстройкой **Outlook**, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="99999-115">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="99999-116">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить в клиенте современную проверку подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="99999-116">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="99999-117">Вы *не* должны полагаться на службу единого входа как на единственный метод проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="99999-117">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="99999-118">Необходимо реализовать альтернативную систему проверки подлинности, к которой ваша надстройка сможет обратиться в случае ошибок.</span><span class="sxs-lookup"><span data-stu-id="99999-118">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="99999-119">Можно использовать систему пользовательских таблиц и проверки подлинности или задействовать одного из поставщиков входа социальных сетей.</span><span class="sxs-lookup"><span data-stu-id="99999-119">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="99999-120">Дополнительные сведения о том, как это сделать с помощью надстройки Office, см. в статье [авторизация внешних служб в надстройке Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins).</span><span class="sxs-lookup"><span data-stu-id="99999-120">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins).</span></span> <span data-ttu-id="99999-121">Для *Outlook* существует рекомендованная альтернативная система.</span><span class="sxs-lookup"><span data-stu-id="99999-121">For *Outlook*, there is a recommended fall back system.</span></span> <span data-ttu-id="99999-122">Дополнительные сведения см. в статье [Сценарий: реализация единого входа для службы в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="99999-122">For more details, see [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="99999-123">Принцип работы единого входа во время выполнения</span><span class="sxs-lookup"><span data-stu-id="99999-123">How it works at runtime</span></span>

<span data-ttu-id="99999-124">На приведенной ниже схеме показано, как работает единый вход.</span><span class="sxs-lookup"><span data-stu-id="99999-124">The following diagram shows how the SSO process works.</span></span>

![Диаграмма, демонстрирующая процесс единого входа](../images/sso-overview-diagram.png)

1. <span data-ttu-id="99999-126">Код JavaScript надстройки вызывает новый API Office.js — [](#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="99999-126">In the add-in, JavaScript calls a new Office.js API [](#sso-api-reference).</span></span> <span data-ttu-id="99999-127">Он указывает ведущему приложению Office, что необходимо получить маркер доступа к надстройке</span><span class="sxs-lookup"><span data-stu-id="99999-127">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="99999-128">См. раздел [Пример маркера доступа](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="99999-128">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="99999-129">Если вход в Office не выполнен, в ведущем приложении открывается всплывающее окно, в котором пользователю предлагается войти.</span><span class="sxs-lookup"><span data-stu-id="99999-129">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="99999-130">Если пользователь запускает надстройку в первый раз, ему предлагается дать согласие.</span><span class="sxs-lookup"><span data-stu-id="99999-130">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="99999-131">Ведущее приложение Office запрашивает **маркер надстройки** у конечной точки Azure AD версии 2.0 для текущего пользователя.</span><span class="sxs-lookup"><span data-stu-id="99999-131">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="99999-132">Azure AD отправляет маркер надстройки ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="99999-132">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="99999-133">Ведущее приложение Office отправляет **маркер** надстройке в составе объекта результата, возвращенного при вызове метода `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="99999-133">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="99999-134">JavaScript в надстройке может проанализировать маркер и извлечь необходимую информацию, например, адрес электронной почты пользователя.</span><span class="sxs-lookup"><span data-stu-id="99999-134">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="99999-135">Кроме того, надстройка может отправить HTTP-запрос на сервер для получения дополнительных сведений о пользователе, например, его настроек.</span><span class="sxs-lookup"><span data-stu-id="99999-135">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="99999-136">Можно также отправить маркер доступа на сервер для анализа и проверки.</span><span class="sxs-lookup"><span data-stu-id="99999-136">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="99999-137">Разработка надстройки с единым входом</span><span class="sxs-lookup"><span data-stu-id="99999-137">Develop an SSO add-in</span></span>

<span data-ttu-id="99999-138">В этом разделе описаны задачи, необходимые для создания надстройки Office с единым входом.</span><span class="sxs-lookup"><span data-stu-id="99999-138">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="99999-139">Эти задачи описываются независимо от языка и платформы.</span><span class="sxs-lookup"><span data-stu-id="99999-139">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="99999-140">Подробные пошаговые инструкции см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="99999-140">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="99999-141">Создание надстройки Office на платформе Node.js с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="99999-141">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="99999-142">Создание надстройки Office на платформе ASP.NET с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="99999-142">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="99999-143">Создание приложения-службы</span><span class="sxs-lookup"><span data-stu-id="99999-143">Create the service application</span></span>

<span data-ttu-id="99999-144">Зарегистрируйте надстройку на портале регистрации конечной точки Azure v2.0: https://apps.dev.microsoft.com.</span><span class="sxs-lookup"><span data-stu-id="99999-144">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span> <span data-ttu-id="99999-145">Этот процесс занимает 5 – 10 минут и включает выполнение следующих задач:</span><span class="sxs-lookup"><span data-stu-id="99999-145">This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="99999-146">Получение идентификатора и секрета клиента для надстройки.</span><span class="sxs-lookup"><span data-stu-id="99999-146">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="99999-147">Укажите разрешения, которые необходимы надстройкам для AAD v.</span><span class="sxs-lookup"><span data-stu-id="99999-147">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="99999-148">2.0 (при необходимости — для Microsoft Graph);</span><span class="sxs-lookup"><span data-stu-id="99999-148">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="99999-149">разрешение "профиля" требуется всегда;</span><span class="sxs-lookup"><span data-stu-id="99999-149">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="99999-150">Предоставление надстройке доверия ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="99999-150">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="99999-151">предварительная авторизация ведущего приложения Office для надстройки с помощью заданного по умолчанию разрешения *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="99999-151">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="99999-152">Для получения дополнительной информации об этом процессе см. статью [Регистрация надстройки Office, использующей единый вход с конечной точкой Azure AD версии 2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="99999-152">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="99999-153">Конфигурация надстройки</span><span class="sxs-lookup"><span data-stu-id="99999-153">Configure the add-in</span></span>

<span data-ttu-id="99999-154">Добавьте новую разметку в манифест надстройки:</span><span class="sxs-lookup"><span data-stu-id="99999-154">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="99999-155">**WebApplicationInfo** — родительский элемент для указанных ниже элементов;</span><span class="sxs-lookup"><span data-stu-id="99999-155">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="99999-156">**Id** — идентификатор клиента надстройки; это идентификатор приложения, который вы получаете в рамках регистрации надстройки;</span><span class="sxs-lookup"><span data-stu-id="99999-156">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="99999-157">См. статью [Регистрация надстройки Office, использующей единый вход с конечной точкой Azure AD версии 2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="99999-157">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="99999-158">**Resource** — URL-адрес надстройки;</span><span class="sxs-lookup"><span data-stu-id="99999-158">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="99999-159">**Scopes** — родительский элемент одного или нескольких элементов **Scope**;</span><span class="sxs-lookup"><span data-stu-id="99999-159">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="99999-160">**Область** — указывает разрешение, необходимое надстройке для работы с AAD.</span><span class="sxs-lookup"><span data-stu-id="99999-160">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="99999-161">Разрешение `profile` требуется всегда, и оно может быть единственным необходимым разрешением, если надстройка не получает доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="99999-161">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="99999-162">Если надстройка получает этот доступ, потребуются элементы **Scope** для необходимых разрешений Microsoft Graph; например, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="99999-162">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="99999-163">Для библиотек, которые используются в коде для доступа к Microsoft Graph, могут потребоваться дополнительные разрешения.</span><span class="sxs-lookup"><span data-stu-id="99999-163">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="99999-164">Например, для библиотеки проверки подлинности Майкрософт (MSAL) для .NET требуется разрешение `offline_access`.</span><span class="sxs-lookup"><span data-stu-id="99999-164">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="99999-165">Для получения дополнительной информации см. статью [Авторизованный доступ в Microsoft Graph из вашей надстройки Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="99999-165">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="99999-p113">Для всех ведущих приложений, кроме Outlook, добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Для Outlook добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="99999-p113">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="99999-168">Ниже приведен пример части кода.</span><span class="sxs-lookup"><span data-stu-id="99999-168">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="99999-169">Добавление кода для клиента</span><span class="sxs-lookup"><span data-stu-id="99999-169">Add client-side code</span></span>

<span data-ttu-id="99999-170">Добавьте в надстройку код JavaScript для:</span><span class="sxs-lookup"><span data-stu-id="99999-170">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="99999-171">Вызов [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="99999-171">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="99999-172">анализа маркера доступа или его передачи в серверный код надстройки;</span><span class="sxs-lookup"><span data-stu-id="99999-172">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="99999-173">Далее представлен простой пример вызова `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="99999-173">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="99999-174">В данном примере явным образом обрабатывается только один тип ошибки.</span><span class="sxs-lookup"><span data-stu-id="99999-174">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="99999-175">Для ознакомления с примерами более сложной обработки ошибок см. статьи [Home.js в Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) и [program.js в Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span><span class="sxs-lookup"><span data-stu-id="99999-175">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="99999-176">См. статью [Сообщения устранения ошибок единого входа (SSO)](troubleshoot-sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="99999-176">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

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

<span data-ttu-id="99999-177">Далее представлен пример передачи маркера надстройки на сервер.</span><span class="sxs-lookup"><span data-stu-id="99999-177">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="99999-178">При отправке запроса обратно на сервер маркер указывается в качестве заголовка `Authorization`.</span><span class="sxs-lookup"><span data-stu-id="99999-178">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="99999-179">Данный пример предусматривает отправку данных JSON, поэтому используется метод `POST`, однако `GET` достаточно для отправки маркера доступа, если не выполняется запись в сервер.</span><span class="sxs-lookup"><span data-stu-id="99999-179">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="99999-180">Когда вызывать метод</span><span class="sxs-lookup"><span data-stu-id="99999-180">When to call the method</span></span>

<span data-ttu-id="99999-181">Если надстройка не может работать без входа в Office, необходимо вызвать `getAccessTokenAsync` *при запуске надстройки*.</span><span class="sxs-lookup"><span data-stu-id="99999-181">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="99999-182">Если в надстройке присутствует функциональность, которая не требует входа пользователя, метод `getAccessTokenAsync` *вызывается тогда, когда пользователь выполняет действие, для которого требуется вход*.</span><span class="sxs-lookup"><span data-stu-id="99999-182">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="99999-183">Нет значительного замедления при повторяющихся вызовах `getAccessTokenAsync`, поскольку Office кэширует маркер доступа и использует его снова, пока не истечет срок его действия, не вызывая конечную точку AAD v.</span><span class="sxs-lookup"><span data-stu-id="99999-183">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="99999-184">2.0 при каждом вызове  `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="99999-184">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="99999-185">Поэтому вызовы `getAccessTokenAsync` можно добавлять во все функции и обработчики, которые инициируют действие, где нужен маркер.</span><span class="sxs-lookup"><span data-stu-id="99999-185">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="99999-186">Добавление серверного кода</span><span class="sxs-lookup"><span data-stu-id="99999-186">Add server-side code</span></span>

<span data-ttu-id="99999-187">В большинстве случаев практически нет смысла получать маркер доступа, если надстройка не передает его на сторону сервера и не использует его там.</span><span class="sxs-lookup"><span data-stu-id="99999-187">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="99999-188">Далее указаны некоторые серверные задачи, которые может выполнять надстройка.</span><span class="sxs-lookup"><span data-stu-id="99999-188">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="99999-189">Создание одного или нескольких методов веб-API, использующих информацию о пользователе, которая извлекается из маркера, например, метод поиска предпочтений пользователя в базе данных на сервере</span><span class="sxs-lookup"><span data-stu-id="99999-189">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="99999-190">(См. статью **Использование маркера единого входа в качестве удостоверения** далее). В зависимости от языка и платформы могут быть доступны библиотеки, который упростят создание нужного кода.</span><span class="sxs-lookup"><span data-stu-id="99999-190">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="99999-191">Получение данных Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="99999-191">Get Microsoft Graph data.</span></span> <span data-ttu-id="99999-192">Серверный код должен:</span><span class="sxs-lookup"><span data-stu-id="99999-192">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="99999-193">проверять маркеры доступа (см. статью **Проверка маркера доступа** далее);</span><span class="sxs-lookup"><span data-stu-id="99999-193">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="99999-194">Инициируйте поток «от имени» с вызовом конечной точки Azure AD версии 2.0, который включает токен доступа, некоторые метаданные о пользователе и учетные данные надстройки (ее идентификатор и секрет).</span><span class="sxs-lookup"><span data-stu-id="99999-194">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="99999-195">в этом контексте маркер доступа называется маркером начальной загрузки;</span><span class="sxs-lookup"><span data-stu-id="99999-195">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="99999-196">выполнять кэширование нового маркера доступа от потока "от имени";</span><span class="sxs-lookup"><span data-stu-id="99999-196">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="99999-197">Получите данные с Microsoft Graph, используя новый маркер.</span><span class="sxs-lookup"><span data-stu-id="99999-197">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="99999-198">Для ознакомления с дополнительной информацией о получении авторизованного доступа к данным пользователя Microsoft Graph см. статью [Авторизованный доступ в Microsoft Graph из вашей надстройки Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="99999-198">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="99999-199">Утвердите маркер доступа</span><span class="sxs-lookup"><span data-stu-id="99999-199">For more information, see Validate the access token.</span></span>

<span data-ttu-id="99999-200">Когда веб-API получит маркер доступа, этот токен необходимо проверить перед использованием.</span><span class="sxs-lookup"><span data-stu-id="99999-200">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="99999-201">Это маркер JSON Web Token (JWT), то есть его проверка выполняется так же, как и в большинстве стандартных потоков OAuth.</span><span class="sxs-lookup"><span data-stu-id="99999-201">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="99999-202">Доступно множество библиотек, которые могут выполнять проверку JWT, но основные действия подразумевают:</span><span class="sxs-lookup"><span data-stu-id="99999-202">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="99999-203">проверку правильности формата маркера;</span><span class="sxs-lookup"><span data-stu-id="99999-203">Checking that the token is well-formed</span></span>
- <span data-ttu-id="99999-204">проверку факта выдачи маркера нужным центром сертификации;</span><span class="sxs-lookup"><span data-stu-id="99999-204">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="99999-205">проверку предназначения маркера для веб-API.</span><span class="sxs-lookup"><span data-stu-id="99999-205">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="99999-206">При проверке маркера следует учитывать приведенные ниже рекомендации.</span><span class="sxs-lookup"><span data-stu-id="99999-206">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="99999-207">Действительные маркеры единого входа выдает центр сертификации Azure, `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="99999-207">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="99999-208">Утверждение `iss` в маркере должно начинаться с этого значения.</span><span class="sxs-lookup"><span data-stu-id="99999-208">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="99999-209">Параметру `aud` маркера будет присвоено значение идентификатора приложения с портала регистрации.</span><span class="sxs-lookup"><span data-stu-id="99999-209">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="99999-210">Для параметра `scp` маркера будет задано значение `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="99999-210">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="99999-211">Использование маркера единого входа в качестве удостоверения</span><span class="sxs-lookup"><span data-stu-id="99999-211">Using the SSO token as an identity</span></span>

<span data-ttu-id="99999-212">Если приложению необходимо проверить удостоверение пользователя, то маркер единого входа содержит сведения, с помощью которых можно такое удостоверение определить.</span><span class="sxs-lookup"><span data-stu-id="99999-212">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="99999-213">Ниже перечислены утверждения из маркера, связанные с удостоверениями.</span><span class="sxs-lookup"><span data-stu-id="99999-213">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="99999-214">`name` — Отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="99999-214">`name` - The user's display name.</span></span>
- <span data-ttu-id="99999-215">`preferred_username` — Адрес электронной почты пользователя.</span><span class="sxs-lookup"><span data-stu-id="99999-215">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="99999-216">`oid` — GUID, предоставляющий ИД пользователя в Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="99999-216">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="99999-217">`tid` — GUID, предоставляющий ИД организации пользователя в Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="99999-217">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="99999-218">Значения `name` и `preferred_username` могут меняться, поэтому рекомендуется использовать значения `oid` и `tid`, чтобы коррелировать удостоверение с внутренней службой авторизации.</span><span class="sxs-lookup"><span data-stu-id="99999-218">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="99999-219">Например, если служба может форматировать эти значения вместе (в виде `{oid-value}@{tid-value}`), то их следует хранить в качестве значения в записи пользователя во внутренней базе данных пользователей.</span><span class="sxs-lookup"><span data-stu-id="99999-219">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="99999-220">При последующих запросах удостоверение пользователя можно будет получать с помощью того же значения, а доступ к определенным ресурсам может предоставляться в соответствии с действующими механизмами управления доступом.</span><span class="sxs-lookup"><span data-stu-id="99999-220">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="99999-221">Пример маркера доступа</span><span class="sxs-lookup"><span data-stu-id="99999-221">Example access token</span></span>

<span data-ttu-id="99999-222">Далее представлены типичные расшифрованные полезные данные маркера доступа.</span><span class="sxs-lookup"><span data-stu-id="99999-222">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="99999-223">Для получения дополнительной информации о свойствах см. статью [Ссылка на маркеры Azure Active Directory v2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="99999-223">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="99999-224">Использование SSO с надстройкой Outlook</span><span class="sxs-lookup"><span data-stu-id="99999-224">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="99999-225">Имеются небольшие, но важные различия между использованием функции единого входа в надстройке Outlook и использованием ее в надстройке Excel, PowerPoint или Word.</span><span class="sxs-lookup"><span data-stu-id="99999-225">There are some small, but important differences in using SSO in and Outlook add-in from using it in as Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="99999-226">Ознакомьтесь со статьями [Аутентификация пользователя с маркером единого входа в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) и [Сценарий: реализация единого входа для службы в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span><span class="sxs-lookup"><span data-stu-id="99999-226">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="99999-227">Справка по API SSO</span><span class="sxs-lookup"><span data-stu-id="99999-227">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="99999-228">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="99999-228">getAccessTokenAsync</span></span>

<span data-ttu-id="99999-229">Проверка подлинности пространства имен Office `Office.context.auth` предоставляет метод `getAccessTokenAsync`, позволяющий основному приложению Office получать маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="99999-229">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="99999-230">Косвенно это также позволяет надстройке получать доступ к данным Microsoft Graph с включенным пользователем, не требуя от пользователя входа во второй раз.</span><span class="sxs-lookup"><span data-stu-id="99999-230">Indirectly, enable the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="99999-231">Этот метод вызывает конечную точку Azure Active Directory версии 2.0, чтобы получить токен доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="99999-231">Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="99999-232">Это позволяет надстройкам идентифицировать пользователей.</span><span class="sxs-lookup"><span data-stu-id="99999-232">This enables add-ins to identify users.</span></span> <span data-ttu-id="99999-233">Код на стороне сервера может использовать этот маркер для доступа к Microsoft Graph, чтобы добавить веб-приложение надстройки с помощью [потока OAuth "от имени пользователя"](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="99999-233">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="99999-234">В Outlook этот интерфейс API не поддерживается, если надстройка загружается в почтовый ящик Outlook.com или Gmail.</span><span class="sxs-lookup"><span data-stu-id="99999-234">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

<table><tr><td><span data-ttu-id="99999-235">Основные приложения</span><span class="sxs-lookup"><span data-stu-id="99999-235">Hosts</span></span></td><td><span data-ttu-id="99999-236">Excel, OneNote, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="99999-236">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="99999-237">Наборы обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="99999-237">Requirement sets</span></span></td><td>[<span data-ttu-id="99999-238">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="99999-238">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="99999-239">Параметры</span><span class="sxs-lookup"><span data-stu-id="99999-239">Parameters</span></span>

<span data-ttu-id="99999-240">`options` — Необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="99999-240">`options` - Optional.</span></span> <span data-ttu-id="99999-241">Принимает `AuthOptions` объект (см. ниже) для определения расширений функциональности входа.</span><span class="sxs-lookup"><span data-stu-id="99999-241">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="99999-242">`callback` — Необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="99999-242">`callback` - Optional.</span></span> <span data-ttu-id="99999-243">Принимает метод обратного вызова, который может анализировать маркер для идентификатора пользователя или использовать маркер в потоке «от имени», чтобы получить доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="99999-243">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="99999-244">Если [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` «успешно завершено», тогда `AsyncResult.value` — это необработанный AAD v.</span><span class="sxs-lookup"><span data-stu-id="99999-244">If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="99999-245">отформатированный маркер доступа 2.0.</span><span class="sxs-lookup"><span data-stu-id="99999-245">2.0-formatted access token.</span></span>

<span data-ttu-id="99999-246">`AuthOptions` Интерфейс предоставляет параметры для взаимодействия с пользователем, когда Office получает маркер доступа для надстройки от AAD v.</span><span class="sxs-lookup"><span data-stu-id="99999-246">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="99999-247">2.0 с методом `getAccessTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="99999-247">2.0 with the `getAccessTokenAsync` method.</span></span>

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



