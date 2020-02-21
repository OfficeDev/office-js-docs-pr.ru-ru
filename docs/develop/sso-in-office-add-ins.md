---
title: Включение единого входа для надстроек Office
description: ''
ms.date: 01/14/2020
localization_priority: Priority
ms.openlocfilehash: 5efb400e0999768abb8ae94e30dd271b2cf69a48
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42162947"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="a5e74-102">Включение единого входа для надстроек Office (тестовый режим)</span><span class="sxs-lookup"><span data-stu-id="a5e74-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="a5e74-103">Пользователи входят в Office (в Интернете, на мобильных устройствах и настольных компьютерах), используя личную учетную запись Майкрософт либо рабочую или учебную учетную запись (Office 365).</span><span class="sxs-lookup"><span data-stu-id="a5e74-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="a5e74-104">Вы можете использовать эту возможность и использовать единый вход для авторизации пользователя в вашей надстройке, при этом пользователю не потребуется входить повторно.</span><span class="sxs-lookup"><span data-stu-id="a5e74-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![Изображение, иллюстрирующее процесс входа в надстройку](../images/sso-for-office-addins.png)

## <a name="preview-status"></a><span data-ttu-id="a5e74-106">Предварительный просмотр состояния</span><span class="sxs-lookup"><span data-stu-id="a5e74-106">Preview Status</span></span>

<span data-ttu-id="a5e74-107">API единого входа в настоящее время поддерживается только для предварительного просмотра.</span><span class="sxs-lookup"><span data-stu-id="a5e74-107">The Single Sign-on API is currently supported in preview only.</span></span> <span data-ttu-id="a5e74-108">Также оно доступно для разработчиков для целей экспериментирования; но его нельзя использовать в рабочей надстройке.</span><span class="sxs-lookup"><span data-stu-id="a5e74-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="a5e74-109">Кроме того, надстройки, которые используют единый вход, не допускаются к использованию в [AppSource](https://appsource.microsoft.com).</span><span class="sxs-lookup"><span data-stu-id="a5e74-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="a5e74-110">Для единого входа требуется Office 365 (версия Office, распространяемая по подписке).</span><span class="sxs-lookup"><span data-stu-id="a5e74-110">SSO requires Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="a5e74-111">Следует использовать последнюю версию для текущего месяца и сборку из канала для участников программы предварительной оценки.</span><span class="sxs-lookup"><span data-stu-id="a5e74-111">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="a5e74-112">Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="a5e74-112">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="a5e74-113">Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="a5e74-113">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="a5e74-114">Обратите внимание на то, что когда сборка будет готова для выпуска на канале Semi-annual channel, поддержка функций предварительного просмотра, включая единый вход, отключается для этой сборки.</span><span class="sxs-lookup"><span data-stu-id="a5e74-114">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

<span data-ttu-id="a5e74-115">Не все приложения Office поддерживают предварительный просмотр единого входа.</span><span class="sxs-lookup"><span data-stu-id="a5e74-115">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="a5e74-116">Эта возможность доступна в Word, Excel, Outlook и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="a5e74-116">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="a5e74-117">Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API идентификации](../reference/requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="a5e74-117">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).</span></span>

## <a name="requirements-and-best-practices"></a><span data-ttu-id="a5e74-118">Рекомендации и требования</span><span class="sxs-lookup"><span data-stu-id="a5e74-118">Requirements and Best Practices</span></span>

> [!NOTE]
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

<span data-ttu-id="a5e74-119">Если вы работаете с надстройкой **Outlook**, обязательно включите современную проверку подлинности для клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="a5e74-119">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="a5e74-120">Сведения о том, как это сделать, см. в статье [Exchange Online: как включить в клиенте современную проверку подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span><span class="sxs-lookup"><span data-stu-id="a5e74-120">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="a5e74-121">Вам *не* следует опираться на функцию единого в качестве единого способа проверки подлинности вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="a5e74-121">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="a5e74-122">Вам следует внедрить альтернативную систему проверки подлинности, на которую ваша надстройка может опираться при определенных ошибках.</span><span class="sxs-lookup"><span data-stu-id="a5e74-122">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="a5e74-123">Вы можете использовать систему таблиц пользователя и проверки подлинности, либо вы можете выделить одну из систем входа с использованием социальных сервисов.</span><span class="sxs-lookup"><span data-stu-id="a5e74-123">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="a5e74-124">Дополнительные сведения о том, как это сделать с помощью надстройки Office см. в статье [Авторизация внешних служб в надстройке Office](auth-external-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="a5e74-124">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](auth-external-add-ins.md).</span></span> <span data-ttu-id="a5e74-125">Для *Outlook* существует рекомендуемая система возврата.</span><span class="sxs-lookup"><span data-stu-id="a5e74-125">For *Outlook*, there is a recommended fallback system.</span></span> <span data-ttu-id="a5e74-126">Дополнительные сведения см. в статье [Сценарий: реализация единого входа для службы в надстройке Outlook](../outlook/implement-sso-in-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="a5e74-126">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](../outlook/implement-sso-in-outlook-add-in.md).</span></span> <span data-ttu-id="a5e74-127">Примеры использования Azure Active Directory в качестве системы возврата см. в статьях [Единый вход с использованием NodeJS для надстройки Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) и [Единый вход с использованием ASP.NET для надстройки Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span><span class="sxs-lookup"><span data-stu-id="a5e74-127">For samples that use Azure Active Directory as the fallback system, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) and [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span></span>

## <a name="how-sso-works-at-runtime"></a><span data-ttu-id="a5e74-128">Принцип работы единого входа во время выполнения</span><span class="sxs-lookup"><span data-stu-id="a5e74-128">How SSO works at runtime</span></span>

<span data-ttu-id="a5e74-129">На приведенной ниже схеме показано, как работает единый вход.</span><span class="sxs-lookup"><span data-stu-id="a5e74-129">The following diagram shows how the SSO process works.</span></span>

![Схема единого входа](../images/sso-overview-diagram.png)

1. <span data-ttu-id="a5e74-131">Код JavaScript надстройки вызывает новый API Office.js — [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span><span class="sxs-lookup"><span data-stu-id="a5e74-131">In the add-in, JavaScript calls a new Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span></span> <span data-ttu-id="a5e74-132">Он указывает ведущему приложению Office, что необходимо получить маркер доступа к надстройке.</span><span class="sxs-lookup"><span data-stu-id="a5e74-132">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="a5e74-133">См. [Пример маркера доступа](#example-access-token).</span><span class="sxs-lookup"><span data-stu-id="a5e74-133">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="a5e74-134">Если вход в Office не выполнен, в ведущем приложении открывается всплывающее окно, в котором пользователю предлагается войти.</span><span class="sxs-lookup"><span data-stu-id="a5e74-134">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="a5e74-135">Если пользователь запускает надстройку в первый раз, ему предлагается дать согласие.</span><span class="sxs-lookup"><span data-stu-id="a5e74-135">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="a5e74-136">Ведущее приложение Office запрашивает **маркер надстройки** у конечной точки Azure AD версии 2.0 для текущего пользователя.</span><span class="sxs-lookup"><span data-stu-id="a5e74-136">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="a5e74-137">Azure AD отправляет маркер надстройки ведущему приложению Office.</span><span class="sxs-lookup"><span data-stu-id="a5e74-137">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="a5e74-138">Ведущее приложение Office отправляет **маркер** надстройке в составе объекта результата, возвращенного при вызове метода `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-138">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessToken` call.</span></span>
7. <span data-ttu-id="a5e74-139">JavaScript в надстройке может проанализировать маркер и извлечь данные, которые необходимы, например адрес электронной почты.</span><span class="sxs-lookup"><span data-stu-id="a5e74-139">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span>
8. <span data-ttu-id="a5e74-140">Кроме того, надстройка может отправить HTTP-запрос на серверную часть для получения дополнительных данных о пользователе, например, настройки пользователя.</span><span class="sxs-lookup"><span data-stu-id="a5e74-140">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="a5e74-141">Кроме того маркер доступа сам может быть отправлен на серверную часть для анализа и проверки.</span><span class="sxs-lookup"><span data-stu-id="a5e74-141">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span>

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="a5e74-142">Разработка надстройки с единым входом</span><span class="sxs-lookup"><span data-stu-id="a5e74-142">Develop an SSO add-in</span></span>

<span data-ttu-id="a5e74-143">В этом разделе описаны задачи, необходимые для создания надстройки Office с единым входом.</span><span class="sxs-lookup"><span data-stu-id="a5e74-143">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="a5e74-144">Эти задачи описываются независимо от языка и платформы.</span><span class="sxs-lookup"><span data-stu-id="a5e74-144">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="a5e74-145">Подробные пошаговые инструкции см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="a5e74-145">For detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="a5e74-146">Создание надстройки Office на платформе Node.js с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="a5e74-146">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="a5e74-147">Создание надстройки Office на платформе ASP.NET с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="a5e74-147">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

> [!NOTE]
> <span data-ttu-id="a5e74-148">Вы можете использовать генератор Yeoman для создания надстройки Office на платформе Node.js с использованием единого входа.</span><span class="sxs-lookup"><span data-stu-id="a5e74-148">You can use the Yeoman generator to create an SSO-enabled, Node.js Office Add-in.</span></span> <span data-ttu-id="a5e74-149">Генератор Yeoman упрощает процесс создания надстройки с использованием единого входа, автоматизируя действия, необходимые для настройки единого входа в Azure, и создавая код, необходимый для его использования в надстройке.</span><span class="sxs-lookup"><span data-stu-id="a5e74-149">The Yeoman generator simplifies the process of creating an SSO-enabled add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="a5e74-150">Дополнительные сведения см. в статье [Краткое руководство по использованию единого входа (SSO)](../quickstarts/sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="a5e74-150">For more information, see the [Single sign-on (SSO) quick start](../quickstarts/sso-quickstart.md).</span></span>

### <a name="create-the-service-application"></a><span data-ttu-id="a5e74-151">Создание приложения-службы</span><span class="sxs-lookup"><span data-stu-id="a5e74-151">Create the service application</span></span>

<span data-ttu-id="a5e74-p111">Зарегистрируйте надстройку на портале регистрации для конечной точки Azure версии 2.0. Этот процесс занимает 5–10 минут и включает следующие задачи:</span><span class="sxs-lookup"><span data-stu-id="a5e74-p111">Register the add-in at the registration portal for the Azure v2.0 endpoint. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="a5e74-154">Получение идентификатора и секрета клиента для надстройки;</span><span class="sxs-lookup"><span data-stu-id="a5e74-154">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="a5e74-155">Указание необходимых надстройке разрешений для конечной точки AAD в.</span><span class="sxs-lookup"><span data-stu-id="a5e74-155">Specify the permissions that your add-in needs to AAD v.</span></span> <span data-ttu-id="a5e74-156">2.0 (и Microsoft Graph в качестве опции).</span><span class="sxs-lookup"><span data-stu-id="a5e74-156">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="a5e74-157">Разрешение «профиль» необходимо во всех случаях.</span><span class="sxs-lookup"><span data-stu-id="a5e74-157">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="a5e74-158">Предоставление надстройке доверия ведущего приложения Office;</span><span class="sxs-lookup"><span data-stu-id="a5e74-158">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="a5e74-159">Предварительная авторизация ведущего приложения Office для надстройки с помощью заданного по умолчанию разрешения *access_as_user*.</span><span class="sxs-lookup"><span data-stu-id="a5e74-159">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="a5e74-160">Дополнительные сведения о данном процессе см. в статье [Регистрация надстройки Office, использующей единый вход с конечной точкой Microsoft Azure AD версии 2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="a5e74-160">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="a5e74-161">Конфигурация надстройки</span><span class="sxs-lookup"><span data-stu-id="a5e74-161">Configure the add-in</span></span>

<span data-ttu-id="a5e74-162">Добавьте новую разметку в манифест надстройки:</span><span class="sxs-lookup"><span data-stu-id="a5e74-162">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="a5e74-163">**WebApplicationInfo** — родительский элемент для указанных ниже элементов.</span><span class="sxs-lookup"><span data-stu-id="a5e74-163">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="a5e74-164">**Id** - Идентификатор клиента надстройки. Это идентификатор приложения, который вы получаете в процессе регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="a5e74-164">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="a5e74-165">См. [Регистрация надстройки Office, использующей единый вход с конечной точкой Microsoft Azure AD версии 2.0](register-sso-add-in-aad-v2.md).</span><span class="sxs-lookup"><span data-stu-id="a5e74-165">See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="a5e74-166">**Resource** — URL-адрес надстройки;</span><span class="sxs-lookup"><span data-stu-id="a5e74-166">**Resource** - The URL of the add-in.</span></span> <span data-ttu-id="a5e74-167">Это тот же URI (включая протокол `api:`), который вы использовали при регистрации надстройки и в AAD.</span><span class="sxs-lookup"><span data-stu-id="a5e74-167">This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD.</span></span> <span data-ttu-id="a5e74-168">Доменная часть данного URI должна соответствовать домену, в том числе поддомену, используемом в URL-адресах в части`<Resources>` манифеста настройки.</span><span class="sxs-lookup"><span data-stu-id="a5e74-168">The domain part of this URI should match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>
* <span data-ttu-id="a5e74-169">**Scopes** — родительский элемент одного или нескольких элементов **Scope**;</span><span class="sxs-lookup"><span data-stu-id="a5e74-169">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="a5e74-170">**Scope** — указывает разрешение, необходимое надстройке для работы с AAD.</span><span class="sxs-lookup"><span data-stu-id="a5e74-170">**Scope** - Specifies a permission that the add-in needs to AAD.</span></span> <span data-ttu-id="a5e74-171">Разрешение `profile` требуется во всех случаях, и оно может быть единственным необходимым разрешением, если ваша надстройка не имеет доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a5e74-171">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="a5e74-172">В противном случае вам также могут потребоваться элементы типа **Область** для необходимым разрешений Microsoft Graph; например, `User.Read`, `Mail.Read`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-172">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="a5e74-173">Библиотеки, которые вы используете в коде, чтобы получить доступ к Microsoft Graph, могут потребовать дополнительные разрешения.</span><span class="sxs-lookup"><span data-stu-id="a5e74-173">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="a5e74-174">Например, библиотека проверки подлинности Microsoft (MSAL) для .NET требует разрешения `offline_access`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-174">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="a5e74-175">Дополнительные сведения см. в статье [Авторизация в Microsoft Graph для надстройки Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="a5e74-175">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="a5e74-p116">Для всех ведущих приложений, кроме Outlook, добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Для Outlook добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-p116">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="a5e74-178">Ниже приведен пример разметки:</span><span class="sxs-lookup"><span data-stu-id="a5e74-178">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="a5e74-179">Добавьте код для клиента</span><span class="sxs-lookup"><span data-stu-id="a5e74-179">Add client-side code</span></span>

<span data-ttu-id="a5e74-180">Добавьте в надстройку код JavaScript для:</span><span class="sxs-lookup"><span data-stu-id="a5e74-180">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="a5e74-181">Вызова [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span><span class="sxs-lookup"><span data-stu-id="a5e74-181">Call [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-).</span></span>

* <span data-ttu-id="a5e74-182">Анализа маркера доступа или передачи его в код надстройки на стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="a5e74-182">Parse the access token or pass it to the add-in’s server-side code.</span></span>

<span data-ttu-id="a5e74-183">Вот простой пример вызова для `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-183">Here's a simple example of a call to `getAccessToken`.</span></span>

> [!NOTE]
> <span data-ttu-id="a5e74-184">В этом примере обрабатывается только один тип ошибки явным образом.</span><span class="sxs-lookup"><span data-stu-id="a5e74-184">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="a5e74-185">Примеры более сложной обработки ошибок см. в статьях [Единый вход с использованием NodeJS для надстройки Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) и [Единый вход с использованием ASP.NET для надстройки Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span><span class="sxs-lookup"><span data-stu-id="a5e74-185">For examples of more elaborate error handling, see [Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO) and [Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO).</span></span>


```js
async function getGraphData() {
    try {
        let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true });

        // The /api/values controller will make the token exchange and use the
        // access token it gets back to make the call to MS Graph.
        getData("/api/DoSomething", bootstrapToken);
    }
    catch (exception) {
        if (exception.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
}
```

<span data-ttu-id="a5e74-186">Вот простой пример передачи маркера надстройки стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="a5e74-186">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="a5e74-187">Маркер включается в виде заголовка `Authorization` при отправке запроса назад стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="a5e74-187">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="a5e74-188">В этом примере представлена отправка данных JSON, поэтому в нем используется метод `POST`, но `GET` достаточно для отправки маркера доступа, когда запись не осуществляется на сервере.</span><span class="sxs-lookup"><span data-stu-id="a5e74-188">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + bootstrapToken
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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="a5e74-189">Когда вызывать метод</span><span class="sxs-lookup"><span data-stu-id="a5e74-189">When to call the method</span></span>

<span data-ttu-id="a5e74-190">Если надстройка не может работать, когда ни один пользователь не выполнил вход в Office, необходимо вызывать метод `getAccessToken` *при запуске надстройки* и передать `allowSignInPrompt: true` в параметр `options` метода `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-190">If your add-in cannot be used when no user is logged into Office, then you should call `getAccessToken` *when the add-in launches* and pass `allowSignInPrompt: true` in the `options` parameter of `getAccessToken`.</span></span>

<span data-ttu-id="a5e74-191">Если ряд функций надстройки может работать без входа пользователя в систему, вы вызываете `getAccessToken` \*, когда пользователь выполняет действие, для которого необходимо выполнить вход \*.</span><span class="sxs-lookup"><span data-stu-id="a5e74-191">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessToken` *when the user takes an action that requires a logged in user*.</span></span> <span data-ttu-id="a5e74-192">Производительность не снижается при повторяющихся вызовах `getAccessToken`, так как Office кэширует маркер начальной загрузки и использует его, пока не истечет срок его действия, не вызывая конечную точку AAD версии</span><span class="sxs-lookup"><span data-stu-id="a5e74-192">There is no significant performance degradation with redundant calls of `getAccessToken` because Office caches the bootstrap token and will reuse it, until it expires, without making another call to the AAD v.</span></span> <span data-ttu-id="a5e74-193">2.0 при каждом вызове `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-193">2.0 endpoint whenever `getAccessToken` is called.</span></span> <span data-ttu-id="a5e74-194">Поэтому вызовы `getAccessToken` можно добавлять во все функции и обработчики, которые инициируют действие, где нужен маркер.</span><span class="sxs-lookup"><span data-stu-id="a5e74-194">So you can add calls of `getAccessToken` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="a5e74-195">Добавление серверного кода</span><span class="sxs-lookup"><span data-stu-id="a5e74-195">Add server-side code</span></span>

<span data-ttu-id="a5e74-196">В большинстве случаев практически нет смысла получать маркер доступа, если надстройка не передает его на сторону сервера и не использует его там.</span><span class="sxs-lookup"><span data-stu-id="a5e74-196">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="a5e74-197">Некоторые задачи на стороне сервера, которые может выполнять ваша надстройка:</span><span class="sxs-lookup"><span data-stu-id="a5e74-197">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="a5e74-198">Создание одного или нескольких методов Web API, которые используют сведения о пользователе, извлеченные из маркера, например, метод, который выполняет поиск параметров пользователя в вашей базе данных.</span><span class="sxs-lookup"><span data-stu-id="a5e74-198">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="a5e74-199">(См. **Использование маркера единого входа в качестве удостоверения** ниже) В зависимости от языка и платформы, могут быть доступны библиотеки, который упростят создание нужного кода.</span><span class="sxs-lookup"><span data-stu-id="a5e74-199">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="a5e74-200">Получение данных Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a5e74-200">Get Microsoft Graph data.</span></span> <span data-ttu-id="a5e74-201">Ваш серверный код должен выполнять следующее:</span><span class="sxs-lookup"><span data-stu-id="a5e74-201">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="a5e74-202">Запуск потока "от имени" с помощью вызова конечной точки Microsoft Azure AD версии 2.0, включающего маркер доступа, некоторые метаданные пользователя и учетные данные надстройки (идентификатор и секрет).</span><span class="sxs-lookup"><span data-stu-id="a5e74-202">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="a5e74-203">В этом контексте маркер доступа вызывается маркером начальной загрузки.</span><span class="sxs-lookup"><span data-stu-id="a5e74-203">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="a5e74-204">Получение данных из Microsoft Graph с помощью нового маркера.</span><span class="sxs-lookup"><span data-stu-id="a5e74-204">Get data from Microsoft Graph by using the new token.</span></span>
    * <span data-ttu-id="a5e74-205">Перед запуском потока можно проверить маркер доступа (см. раздел **Проверка маркера доступа** ниже).</span><span class="sxs-lookup"><span data-stu-id="a5e74-205">Optionally, before initiating the flow, validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="a5e74-206">После завершения потока "от имени" можно кэшировать новый маркер доступа, полученный из этого потока, чтобы использовать его повторно в других вызовах к Microsoft Graph до истечения его срока действия.</span><span class="sxs-lookup"><span data-stu-id="a5e74-206">Optionally, after the on-behalf-of flow completes, cache the new access token that is returned from the flow so that it an be reused in other calls to Microsoft Graph until it expires.</span></span>

 <span data-ttu-id="a5e74-207">Дополнительные сведения о получении авторизованного доступа к данным Microsoft Graph пользователя см. в статье [Авторизация в Microsoft Graph для надстройки Office](authorize-to-microsoft-graph.md).</span><span class="sxs-lookup"><span data-stu-id="a5e74-207">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="a5e74-208">Проверка маркера доступа</span><span class="sxs-lookup"><span data-stu-id="a5e74-208">Validate the access token</span></span>

<span data-ttu-id="a5e74-209">Когда веб-API получит маркер доступа, этот маркер можно проверить перед использованием.</span><span class="sxs-lookup"><span data-stu-id="a5e74-209">Once the Web API receives the access token, it can validate it before using it.</span></span> <span data-ttu-id="a5e74-210">Это маркер JSON Web Token (JWT), то есть его проверка выполняется так же, как и в большинстве стандартных потоков OAuth.</span><span class="sxs-lookup"><span data-stu-id="a5e74-210">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="a5e74-211">Доступно множество библиотек, которые могут выполнять проверку JWT, основные их действия:</span><span class="sxs-lookup"><span data-stu-id="a5e74-211">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="a5e74-212">проверяют правильность формата маркера;</span><span class="sxs-lookup"><span data-stu-id="a5e74-212">Checking that the token is well-formed</span></span>
- <span data-ttu-id="a5e74-213">проверяют, выдан ли маркер нужным центром сертификации;</span><span class="sxs-lookup"><span data-stu-id="a5e74-213">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="a5e74-214">проверяют, предназначен ли маркер для веб-API.</span><span class="sxs-lookup"><span data-stu-id="a5e74-214">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="a5e74-215">При проверке маркера учитывайте приведенные ниже рекомендации.</span><span class="sxs-lookup"><span data-stu-id="a5e74-215">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="a5e74-216">Действительные маркеры единого входа выдает центр сертификации Azure, `https://login.microsoftonline.com`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-216">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="a5e74-217">Утверждение `iss` в маркере должно начинаться с этого значения.</span><span class="sxs-lookup"><span data-stu-id="a5e74-217">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="a5e74-218">Параметру `aud` маркера будет присвоено значение идентификатора приложения с портала регистрации.</span><span class="sxs-lookup"><span data-stu-id="a5e74-218">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="a5e74-219">Для параметра `scp` маркера будет задано значение `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-219">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="a5e74-220">Использование маркера единого входа в качестве удостоверения</span><span class="sxs-lookup"><span data-stu-id="a5e74-220">Using the SSO token as an identity</span></span>

<span data-ttu-id="a5e74-221">Если приложению необходимо проверить удостоверение пользователя, то маркер единого входа содержит сведения, с помощью которых можно определить его.</span><span class="sxs-lookup"><span data-stu-id="a5e74-221">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="a5e74-222">Ниже перечислены утверждения из маркера, связанные с удостоверениями.</span><span class="sxs-lookup"><span data-stu-id="a5e74-222">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="a5e74-223">`name` — Отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="a5e74-223">`name` - The user's display name.</span></span>
- <span data-ttu-id="a5e74-224">`preferred_username` — Электронный адрес пользователя.</span><span class="sxs-lookup"><span data-stu-id="a5e74-224">`preferred_username` - The user's email address.</span></span>
- <span data-ttu-id="a5e74-225">`oid` — GUID, предоставляющий ИД пользователя в Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="a5e74-225">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="a5e74-226">`tid` — GUID, предоставляющий ИД организации пользователя в Azure Active Directory.</span><span class="sxs-lookup"><span data-stu-id="a5e74-226">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="a5e74-227">Значения `name` и `preferred_username` могут меняться, мы рекомендуем использовать значения `oid` и `tid`, чтобы связать удостоверение с внутренней службой авторизации.</span><span class="sxs-lookup"><span data-stu-id="a5e74-227">Since the `name` and `preferred_username` values could change, we recommend that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="a5e74-228">Например, если служба может форматировать эти значения вместе (в виде `{oid-value}@{tid-value}`), то их следует хранить в качестве значения в записи пользователя во внутренней базе данных пользователей.</span><span class="sxs-lookup"><span data-stu-id="a5e74-228">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="a5e74-229">При последующих запросах удостоверение пользователя можно будет получать с помощью того же значения, а доступ к определенным ресурсам может предоставляться в соответствии с действующими механизмами управления доступом.</span><span class="sxs-lookup"><span data-stu-id="a5e74-229">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="a5e74-230">Пример маркера доступа</span><span class="sxs-lookup"><span data-stu-id="a5e74-230">Example access token</span></span>

<span data-ttu-id="a5e74-231">Ниже приведен типичная раскодированная нагрузка маркера доступа.</span><span class="sxs-lookup"><span data-stu-id="a5e74-231">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="a5e74-232">Сведения о свойствах см. в статье [Справочные материалы для маркеров в Azure Active Directory 2.0](/azure/active-directory/develop/active-directory-v2-tokens).</span><span class="sxs-lookup"><span data-stu-id="a5e74-232">For information about the properties, see [Azure Active Directory v2.0 tokens reference](/azure/active-directory/develop/active-directory-v2-tokens).</span></span>

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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="a5e74-233">С использованием единого входа с надстройкой Outlook</span><span class="sxs-lookup"><span data-stu-id="a5e74-233">Using SSO with an Outlook add-in</span></span>

<span data-ttu-id="a5e74-234">Существует ряд небольшие, но важных различий при использовании единого входа в надстройке Outlook и его использования в надстройках Excel, PowerPoint или Word.</span><span class="sxs-lookup"><span data-stu-id="a5e74-234">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="a5e74-235">Обязательно ознакомьтесь с содержанием статей [Выполнение проверки подлинности пользователя с маркером единого входа в надстройке Outlook](../outlook/authenticate-a-user-with-an-sso-token.md) и [Сценарий: Реализация единого входа для вашей службы в надстройке Outlook](../outlook/implement-sso-in-outlook-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="a5e74-235">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](../outlook/authenticate-a-user-with-an-sso-token.md) and [Scenario: Implement single sign-on to your service in an Outlook add-in](../outlook/implement-sso-in-outlook-add-in.md).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="a5e74-236">Справочные материалы по API единого входа</span><span class="sxs-lookup"><span data-stu-id="a5e74-236">SSO API reference</span></span>

### <a name="getaccesstoken"></a><span data-ttu-id="a5e74-237">getAccessToken</span><span class="sxs-lookup"><span data-stu-id="a5e74-237">getAccessToken</span></span>

<span data-ttu-id="a5e74-238">В пространстве имен OfficeRuntime [Auth](/javascript/api/office-runtime/officeruntime.auth) (`OfficeRuntime.Auth`) имеется метод `getAccessToken`, позволяющий узлу Office получать маркер доступа для веб-приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="a5e74-238">The OfficeRuntime [Auth](/javascript/api/office-runtime/officeruntime.auth) namespace, `OfficeRuntime.Auth`, provides a method, `getAccessToken` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="a5e74-239">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="a5e74-239">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessToken(options?: AuthOptions: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="a5e74-240">Метод вызывает конечную точку Azure Active Directory версии 2.0, чтобы получить маркер доступа к вашей надстройке в веб-приложении.</span><span class="sxs-lookup"><span data-stu-id="a5e74-240">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="a5e74-241">Это позволяет надстройкам идентифицировать пользователей.</span><span class="sxs-lookup"><span data-stu-id="a5e74-241">This enables add-ins to identify users.</span></span> <span data-ttu-id="a5e74-242">Код на стороне сервера может использовать этот маркер для доступа к Microsoft Graph, чтобы добавить веб-приложение надстройки с помощью [потока OAuth "от имени пользователя"](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span><span class="sxs-lookup"><span data-stu-id="a5e74-242">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="a5e74-243">В Outlook эта API не поддерживается, если надстройка загружается в почтовый ящик Outlook.com или Gmail.</span><span class="sxs-lookup"><span data-stu-id="a5e74-243">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

|<span data-ttu-id="a5e74-244">Узлы</span><span class="sxs-lookup"><span data-stu-id="a5e74-244">Hosts</span></span>|<span data-ttu-id="a5e74-245">Excel, OneNote, Outlook, PowerPoint и Word.</span><span class="sxs-lookup"><span data-stu-id="a5e74-245">Excel, OneNote, Outlook, PowerPoint, Word</span></span>|
|---|---|
|[<span data-ttu-id="a5e74-246">Наборы требований</span><span class="sxs-lookup"><span data-stu-id="a5e74-246">Requirement sets</span></span>](specify-office-hosts-and-api-requirements.md)|[<span data-ttu-id="a5e74-247">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="a5e74-247">IdentityAPI</span></span>](../reference/requirement-sets/identity-api-requirement-sets.md)|

#### <a name="parameters"></a><span data-ttu-id="a5e74-248">Параметры</span><span class="sxs-lookup"><span data-stu-id="a5e74-248">Parameters</span></span>

<span data-ttu-id="a5e74-249">`options` - Опционально.</span><span class="sxs-lookup"><span data-stu-id="a5e74-249">`options` - Optional.</span></span> <span data-ttu-id="a5e74-250">Принимает объект [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) (см. ниже) для определения поведения при входе.</span><span class="sxs-lookup"><span data-stu-id="a5e74-250">Accepts an [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="a5e74-251">`callback` - Опционально.</span><span class="sxs-lookup"><span data-stu-id="a5e74-251">`callback` - Optional.</span></span> <span data-ttu-id="a5e74-252">Принимает метод обратного вызова, который может выполнить анализ маркера для идентификатора пользователя или использовать маркер в потоке «от имени ваших», чтобы получать доступ к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a5e74-252">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="a5e74-253">Если [AsyncResult](/javascript/api/office/office.asyncresult) `.status` был выполнен «успешно», тогда `AsyncResult.value` представляет собой необработанный маркер доступа AAD</span><span class="sxs-lookup"><span data-stu-id="a5e74-253">If [AsyncResult](/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="a5e74-254">версии 2.0.</span><span class="sxs-lookup"><span data-stu-id="a5e74-254">2.0-formatted access token.</span></span>

<span data-ttu-id="a5e74-255">Интерфейс [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) предоставляет опции для взаимодействия с пользователем, когда Office получает маркер доступа для надстройки из AAD версии</span><span class="sxs-lookup"><span data-stu-id="a5e74-255">The [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="a5e74-256">2.0 с методом `getAccessToken`.</span><span class="sxs-lookup"><span data-stu-id="a5e74-256">2.0 with the `getAccessToken` method.</span></span>
