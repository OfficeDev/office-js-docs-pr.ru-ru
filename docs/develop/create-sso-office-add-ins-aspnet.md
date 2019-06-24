---
title: Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET
description: ''
ms.date: 04/15/2019
localization_priority: Priority
ms.openlocfilehash: a28178fb309450f59435d678c013a7a73bb60978
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128164"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="a8328-102">Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="a8328-102">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="a8328-p101">Ваша надстройка может предоставлять пользователям доступ к нескольким приложениям, используя учетные данные, введенные при входе в Office. [Как включить единый вход в надстройке Office](sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="a8328-p101">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="a8328-105">Из этой статьи вы узнаете, как включить единый вход в надстройке, созданной с помощью ASP.NET, OWIN и MSAL для .NET.</span><span class="sxs-lookup"><span data-stu-id="a8328-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET.</span></span>

> [!NOTE]
> <span data-ttu-id="a8328-106">Сведения о создании надстройки, в которой используется единый вход, на основе Node.js см. в [этой статье](create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="a8328-106">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a8328-107">Предварительные условия</span><span class="sxs-lookup"><span data-stu-id="a8328-107">Prerequisites</span></span>

* <span data-ttu-id="a8328-108">Последняя доступная версия Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="a8328-108">The latest available version of Visual Studio 2017.</span></span>

* <span data-ttu-id="a8328-109">Office 365 (версии Office, распространяемые по подписке).</span><span class="sxs-lookup"><span data-stu-id="a8328-109">Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="a8328-110">Последняя версия для текущего месяца и сборка из канала для участников программы предварительной оценки.</span><span class="sxs-lookup"><span data-stu-id="a8328-110">Latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="a8328-111">Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="a8328-111">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="a8328-112">Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="a8328-112">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="a8328-113">Обратите внимание на то, что когда сборка будет готова для выпуска на канале Semi-annual channel, поддержка функций предварительного просмотра, включая единый вход, отключается для этой сборки.</span><span class="sxs-lookup"><span data-stu-id="a8328-113">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="a8328-114">Настройка начального проекта</span><span class="sxs-lookup"><span data-stu-id="a8328-114">Set up the starter project</span></span>

1. <span data-ttu-id="a8328-115">Клонируйте или скачайте репозиторий [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span><span class="sxs-lookup"><span data-stu-id="a8328-115">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

1. <span data-ttu-id="a8328-p103">Перейдите в папку **Before** и откройте SLN-файл в Visual Studio. Это начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы.</span><span class="sxs-lookup"><span data-stu-id="a8328-p103">Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a8328-p104">В том же репозитории есть готовая версия примера. Она идентична надстройке, которую вы создадите, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью. Чтобы использовать готовую версию, просто откройте файл `sln` и выполните действия, описанные в этой статье, пропустив разделы **Код на стороне клиента** и **Код на стороне сервера**.</span><span class="sxs-lookup"><span data-stu-id="a8328-p104">There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures in this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the `sln` file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side.</span></span>

1. <span data-ttu-id="a8328-p105">Открыв проект, выполните его сборку в Visual Studio. При этом будут установлены пакеты, указанные в файле packages.config. Это может занять от пары секунд до нескольких минут в зависимости от того, сколько пакетов хранится в локальном кэше пакетов на компьютере.</span><span class="sxs-lookup"><span data-stu-id="a8328-p105">After the project opens, build it in Visual Studio, which will install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a8328-p106">Вы увидите сообщение об ошибке, касающейся пространства имен Identity. Это побочный эффект проблемы с конфигурацией, которую вы устраните на следующем этапе. Важно то, что пакеты устанавливаются.</span><span class="sxs-lookup"><span data-stu-id="a8328-p106">You will get an error about the Identity namespace. This is a side effect of a configuration issue that you will fix with the next step. The important thing is that the packages are installed.</span></span>

1. <span data-ttu-id="a8328-127">В настоящий момент версия библиотеки MSAL (Microsoft.Identity.Client), которая нужна для единого входа (версия `1.1.4-preview0002`), не включена в стандартный каталог NuGet, поэтому не указана в package.config. Ее нужно установить отдельно.</span><span class="sxs-lookup"><span data-stu-id="a8328-127">Currently, the version of the MSAL library (Microsoft.Identity.Client) that you need for SSO (version `1.1.4-preview0002`) is not part of the standard nuget catalog, so it is not listed in the package.config, and it must be installed separately.</span></span>

   > 1. <span data-ttu-id="a8328-128">В меню **Сервис** выберите **Диспетчер пакетов NuGet** > **Консоль диспетчера пакетов**.</span><span class="sxs-lookup"><span data-stu-id="a8328-128">On the **Tools** menu, navigate to **Nuget Package Manager** > **Package Manager Console**.</span></span>
   > 2. <span data-ttu-id="a8328-129">В консоли выполните указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="a8328-129">At the console, run the following command.</span></span> <span data-ttu-id="a8328-130">Выполнение может занять минуту или больше времени, даже при быстром подключении к Интернету.</span><span class="sxs-lookup"><span data-stu-id="a8328-130">It may take a minute or more to complete even with a fast Internet connection.</span></span> <span data-ttu-id="a8328-131">Когда все будет готово, в нижней части окна консоли отобразится такое сообщение: **"Microsoft.Identity.Client 1.1.4-preview0002" успешно установлено...**.</span><span class="sxs-lookup"><span data-stu-id="a8328-131">When it finishes you should see **Successfully installed 'Microsoft.Identity.Client 1.1.4-preview0002' ...** near the end of the output in the console.</span></span>
   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`
   > 3. <span data-ttu-id="a8328-132">В **обозревателе решений** разверните элемент **Ссылки** проекта **Office-Add-in-ASPNET-SSO-WebAPI**.</span><span class="sxs-lookup"><span data-stu-id="a8328-132">In **Solution Explorer**, expand **References** of **Office-Add-in-ASPNET-SSO-WebAPI** project.</span></span> <span data-ttu-id="a8328-133">Убедитесь, что в него включена библиотека **Microsoft.Identity.Client**.</span><span class="sxs-lookup"><span data-stu-id="a8328-133">Verify that **Microsoft.Identity.Client** is listed.</span></span> <span data-ttu-id="a8328-134">Если ее нет или она есть, но рядом с нею отображается значок предупреждения, удалите эту запись, а затем с помощью мастера добавления ссылок Visual Studio добавьте ссылку в сборку, указав **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**</span><span class="sxs-lookup"><span data-stu-id="a8328-134">If it is not or there is a warning icon on its entry, delete the entry and then use the Visual Studio Add Reference Wizard to add a reference to the assembly at **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**</span></span>

1. <span data-ttu-id="a8328-135">Еще раз выполните сборку проекта.</span><span class="sxs-lookup"><span data-stu-id="a8328-135">Build the project a second time.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="a8328-136">Регистрация надстройки в конечной точке Azure AD версии 2.0</span><span class="sxs-lookup"><span data-stu-id="a8328-136">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="a8328-137">Следующие инструкции содержат общую информацию, поэтому их можно использовать в нескольких местах.</span><span class="sxs-lookup"><span data-stu-id="a8328-137">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="a8328-138">В рамках этой статьи сделайте вот что:</span><span class="sxs-lookup"><span data-stu-id="a8328-138">For this article do the following:</span></span>

- <span data-ttu-id="a8328-139">Замените заполнитель **$ADD-IN-NAME$** на `Office-Add-in-ASPNET-SSO`.</span><span class="sxs-lookup"><span data-stu-id="a8328-139">Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-ASPNET-SSO`.</span></span>
- <span data-ttu-id="a8328-140">Замените заполнитель **$FQDN-WITHOUT-PROTOCOL$** на `localhost:44355`.</span><span class="sxs-lookup"><span data-stu-id="a8328-140">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:44355`.</span></span>
- <span data-ttu-id="a8328-141">Указывая разрешения в диалоговом окне **Выбор разрешений**, установите флажки для приведенных ниже разрешений.</span><span class="sxs-lookup"><span data-stu-id="a8328-141">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="a8328-142">Для самой надстройки требуется только первое разрешение, а `offline_access` и `openid` требуются для библиотеки MSAL, используемой кодом на стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="a8328-142">Only the first is really required by your add-in itself; but the MSAL library that the server-side code uses requires `offline_access` and `openid`.</span></span> <span data-ttu-id="a8328-143">Разрешение `profile` необходимо, чтобы ведущее приложение Office получило токен для веб-приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8328-143">The `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
  * <span data-ttu-id="a8328-144">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="a8328-144">Files.Read.All</span></span>
  * <span data-ttu-id="a8328-145">offline_access</span><span class="sxs-lookup"><span data-stu-id="a8328-145">offline_access</span></span>
  * <span data-ttu-id="a8328-146">openid</span><span class="sxs-lookup"><span data-stu-id="a8328-146">openid</span></span>
  * <span data-ttu-id="a8328-147">profile</span><span class="sxs-lookup"><span data-stu-id="a8328-147">profile</span></span>


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="a8328-148">Предоставление надстройке разрешений администратора</span><span class="sxs-lookup"><span data-stu-id="a8328-148">Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="a8328-149">Конфигурация надстройки</span><span class="sxs-lookup"><span data-stu-id="a8328-149">Configure the add-in</span></span>

1. <span data-ttu-id="a8328-150">В приведенной ниже строке замените заполнитель {tenant_ID} на идентификатор клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="a8328-150">In the following string, replace the placeholder “{tenant_ID}” with your Office 365 tenancy ID.</span></span> <span data-ttu-id="a8328-151">Если вы не скопировали идентификатор клиента при регистрации надстройки с помощью AAD, воспользуйтесь одним из способов, описанных в статье [Поиск идентификатора клиента Office 365](/onedrive/find-your-office-365-tenant-id).</span><span class="sxs-lookup"><span data-stu-id="a8328-151">If you didn't copy the tenancy ID when you registered the add-in with AAD, use one of the methods in [Find your Office 365 tenant ID](/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span>

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. <span data-ttu-id="a8328-152">В Visual Studio откройте файл web.config. В разделе **appSettings** есть ключи, которым необходимо назначить значения.</span><span class="sxs-lookup"><span data-stu-id="a8328-152">In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.</span></span>

1. <span data-ttu-id="a8328-p112">Используйте строку, составленную на шаге 1, в качестве значения ключа ida:Issuer. Убедитесь, что в значении нет пробелов.</span><span class="sxs-lookup"><span data-stu-id="a8328-p112">Use the string you constructed in step 1 as the value to the key named “ida:Issuer”. Be sure there are no blank spaces in the value.</span></span>

1. <span data-ttu-id="a8328-155">Введите указанные ниже значения для соответствующих ключей.</span><span class="sxs-lookup"><span data-stu-id="a8328-155">Assign the following values to the corresponding keys:</span></span>

    |<span data-ttu-id="a8328-156">Ключ</span><span class="sxs-lookup"><span data-stu-id="a8328-156">Key</span></span>|<span data-ttu-id="a8328-157">Значение</span><span class="sxs-lookup"><span data-stu-id="a8328-157">Value</span></span>|
    |:-----|:-----|
    |<span data-ttu-id="a8328-158">ida:ClientID</span><span class="sxs-lookup"><span data-stu-id="a8328-158">ida:ClientID</span></span>|<span data-ttu-id="a8328-159">Идентификатор приложения, полученный во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8328-159">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="a8328-160">ida:Audience</span><span class="sxs-lookup"><span data-stu-id="a8328-160">ida:Audience</span></span>|<span data-ttu-id="a8328-161">Идентификатор приложения, полученный во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8328-161">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="a8328-162">ida:Password</span><span class="sxs-lookup"><span data-stu-id="a8328-162">ida:Password</span></span>|<span data-ttu-id="a8328-163">Пароль, который вы получили во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8328-163">The password you obtained when you registered the add-in.</span></span>|

   <span data-ttu-id="a8328-p113">Ниже показан пример того, как должны выглядеть четыре измененные вами ключа. *Обратите внимание, что параметры ClientID и Audience имеют одинаковые значения*. Вы также можете использовать один ключ для обеих целей, но вашу разметку web.config будет проще повторно использовать, если вы разделите их, так как они не всегда будут одинаковыми. Кроме того, наличие отдельных ключей позволяет считать вашу надстройку и ресурсом OAuth, связанным с ведущим приложением Office, и клиентом OAuth, связанным с Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a8328-p113">The following is an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*. You can also use a single key for both purposes, but your web.config markup is more reusable if you keep them separate because they aren't always the same. Also, having separate keys reinforces the idea that your add-in is both an OAuth resource, relative to the Office host, and an OAuth client, relative to Microsoft Graph.</span></span>

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />

    ```

   > [!NOTE]
   > <span data-ttu-id="a8328-168">Оставьте остальные параметры в разделе **appSettings** без изменений.</span><span class="sxs-lookup"><span data-stu-id="a8328-168">Leave the other settings in the **appSettings** section unchanged.</span></span>

1. <span data-ttu-id="a8328-169">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="a8328-169">Save and close the file.</span></span>

1. <span data-ttu-id="a8328-170">В проекте надстройки откройте файл манифеста Office-Add-in-ASPNET-SSO.xml.</span><span class="sxs-lookup"><span data-stu-id="a8328-170">In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml”.</span></span>

1. <span data-ttu-id="a8328-171">Перейдите в конец кода файла.</span><span class="sxs-lookup"><span data-stu-id="a8328-171">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="a8328-172">Над закрывающим тегом `</VersionOverrides>` вы найдете следующую часть кода:</span><span class="sxs-lookup"><span data-stu-id="a8328-172">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="a8328-173">Замените заполнитель {application_GUID here} *в обоих местах* разметки идентификатором приложения, скопированным во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8328-173">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="a8328-174">Символы "{}" не входят в состав идентификатора, их не нужно вставлять.</span><span class="sxs-lookup"><span data-stu-id="a8328-174">The "{}" are not part of the ID, so do not include them.</span></span> <span data-ttu-id="a8328-175">Это тот же идентификатор, который использовался для ClientID и Audience в файле web.config.</span><span class="sxs-lookup"><span data-stu-id="a8328-175">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="a8328-176">Значение **Resource** представляет собой **универсальный код ресурса (URI) идентификатора приложения**, который вы задали, когда добавляли платформу веб-API при регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8328-176">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="a8328-177">Раздел **Scopes** используется для создания диалогового окна предоставления разрешений, только если надстройка продается в AppSource.</span><span class="sxs-lookup"><span data-stu-id="a8328-177">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="a8328-178">Откройте вкладку **Предупреждения** в **списке ошибок** в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="a8328-178">Open the **Warnings** tab of the **Error List** in Visual Studio.</span></span> <span data-ttu-id="a8328-179">Если на ней есть предупреждение о том, что `<WebApplicationInfo>` не является допустимым дочерним элементом узла `<VersionOverrides>`, это означает, что используемой вами версии Visual Studio 2017 Preview не удается распознать разметку единого входа.</span><span class="sxs-lookup"><span data-stu-id="a8328-179">If there is a warning that `<WebApplicationInfo>` is not a valid child of `<VersionOverrides>`, your version of Visual Studio 2017 Preview does not recognize the SSO markup.</span></span> <span data-ttu-id="a8328-180">В качестве обходного решения в надстройке Word, Excel или PowerPoint можно выполнить указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="a8328-180">As a workaround, do the following for a Word, Excel, or PowerPoint add-in.</span></span> <span data-ttu-id="a8328-181">Если вы работаете с надстройкой Outlook, вы найдете решение ниже.</span><span class="sxs-lookup"><span data-stu-id="a8328-181">(If you are working with an Outlook add-in see the workaround below.)</span></span>

   - <span data-ttu-id="a8328-182">**Обходное решение для Word, Excel и Powerpoint**</span><span class="sxs-lookup"><span data-stu-id="a8328-182">**Workaround for Word, Excel, and PowerPoint**</span></span>

        1. <span data-ttu-id="a8328-183">Закомментируйте раздел `<WebApplicationInfo>` в манифесте прямо перед завершением узла `</VersionOverrides>`.</span><span class="sxs-lookup"><span data-stu-id="a8328-183">Comment out the `<WebApplicationInfo>` section from the manifest just above the end of `</VersionOverrides>`.</span></span>

        2. <span data-ttu-id="a8328-p116">Нажмите клавишу **F5**, чтобы запустить сеанс отладки. В результате будет создана копия манифеста в следующей папке (доступ к которой проще получить в **проводнике**, чем в Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span><span class="sxs-lookup"><span data-stu-id="a8328-p116">Press **F5** to start a debugging session. This will create a copy of the manifest in the following folder (which is easier to access in **File Explorer** than in Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span></span>

        3. <span data-ttu-id="a8328-186">В копии манифеста удалите синтаксис комментария для раздела `<WebApplicationInfo>`.</span><span class="sxs-lookup"><span data-stu-id="a8328-186">In the copy of the manifest, remove the comment syntax around the `<WebApplicationInfo>` section.</span></span>

        4. <span data-ttu-id="a8328-187">Сохраните копию манифеста.</span><span class="sxs-lookup"><span data-stu-id="a8328-187">Save the copy of the manifest.</span></span>

        5. <span data-ttu-id="a8328-p117">Теперь необходимо принять меры, чтобы Visual Studio не перезаписал копию манифеста, когда вы в следующий раз нажмете клавишу F5. Щелкните правой кнопкой мыши узел решения в верхней части **обозревателя решений** (но не узлы проектов).</span><span class="sxs-lookup"><span data-stu-id="a8328-p117">Now you must prevent Visual Studio from overwriting the copy of the manifest the next time you press F5. Right-click the solution node at the very top of **Solution Explorer** (not either of the project nodes).</span></span>

        6. <span data-ttu-id="a8328-190">В контекстном меню выберите **Свойства**. Откроется диалоговое окно **Страницы свойств решения**.</span><span class="sxs-lookup"><span data-stu-id="a8328-190">Select **Properties** from the context menu and a **Solution Property Pages** dialog box opens.</span></span>

        7. <span data-ttu-id="a8328-191">Разверните пункт **Свойства конфигурации** и щелкните **Конфигурация**.</span><span class="sxs-lookup"><span data-stu-id="a8328-191">Expand **Configuration Properties** and select **Configuration**.</span></span>

        8. <span data-ttu-id="a8328-192">Снимите флажки **Выполнить сборку** и **Развернуть** в строке для проекта **Office-Add-in-ASPNET-SSO** (но *не* проекта **Office-Add-in-ASPNET-SSO-WebAPI**).</span><span class="sxs-lookup"><span data-stu-id="a8328-192">Deselect **Build** and **Deploy** in the row for the **Office-Add-in-ASPNET-SSO** project (*not* the **Office-Add-in-ASPNET-SSO-WebAPI** project).</span></span>

        9. <span data-ttu-id="a8328-193">Закройте диалоговое окно, нажав кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="a8328-193">Press **OK** to close the dialog box.</span></span>

   - <span data-ttu-id="a8328-194">**Обходное решение для Outlook**</span><span class="sxs-lookup"><span data-stu-id="a8328-194">**Workaround for Outlook**</span></span>

        1. <span data-ttu-id="a8328-p118">Найдите файл `MailAppVersionOverridesV1_1.xsd` на компьютере, используемом для разработки. Он должен находиться в том каталоге, в котором установлена среда Visual Studio, в папке `./Xml/Schemas/{lcid}`. Например, при обычной установке 32-разрядной версии VS 2017 в системе, где используется английский язык (США), полный путь будет выглядеть так: `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span><span class="sxs-lookup"><span data-stu-id="a8328-p118">On your development machine, locate the existing `MailAppVersionOverridesV1_1.xsd`. This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`. For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span></span>

        2. <span data-ttu-id="a8328-198">Измените имя существующего файла на `MailAppVersionOverridesV1_1.old`.</span><span class="sxs-lookup"><span data-stu-id="a8328-198">Rename the existing file to `MailAppVersionOverridesV1_1.old`.</span></span>

        3. <span data-ttu-id="a8328-199">Скопируйте измененную версию файла в папку: [Измененная схема MailAppVersionOverrides](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/master/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span><span class="sxs-lookup"><span data-stu-id="a8328-199">Copy this modified version of the file into the folder: [Modified MailAppVersionOverrides Schema](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/master/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span></span>

1. <span data-ttu-id="a8328-200">Сохраните и закройте главный файл манифеста в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="a8328-200">Save and close the main manifest file in Visual Studio.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="a8328-201">Код на стороне клиента</span><span class="sxs-lookup"><span data-stu-id="a8328-201">Code the client side</span></span>

1. <span data-ttu-id="a8328-p119">Откройте файл Home.js в папке **Scripts**. В нем уже есть следующий код:</span><span class="sxs-lookup"><span data-stu-id="a8328-p119">Open the Home.js file in the **Scripts** folder. It already has some code in it:</span></span>
    * <span data-ttu-id="a8328-204">Назначение методу `Office.initialize`, которое, в свою очередь, назначает обработчик события для нажатия кнопки `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="a8328-204">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="a8328-205">Метод `showResult` для отображения сообщения об ошибке (или данных, возвращаемых из Microsoft Graph) в нижней части области задач.</span><span class="sxs-lookup"><span data-stu-id="a8328-205">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="a8328-206">Метод `logErrors` для регистрации в консоли ошибок, которые не предназначены для пользователя.</span><span class="sxs-lookup"><span data-stu-id="a8328-206">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

1. <span data-ttu-id="a8328-p120">После назначения для метода `Office.initialize` добавьте приведенный ниже код. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p120">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="a8328-p121">При обработке ошибок в надстройке иногда автоматически выполняется еще одна попытка получить маркер доступа с помощью другого набора параметров. Переменная счетчика `timesGetOneDriveFilesHasRun` и переменная флажка `triedWithoutForceConsent` используются, чтобы предотвратить циклическое повторение неудачных попыток получить маркер.</span><span class="sxs-lookup"><span data-stu-id="a8328-p121">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options. The counter variable `timesGetOneDriveFilesHasRun`, and the flag variable `triedWithoutForceConsent` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="a8328-p122">Метод `getDataWithToken` создается на следующем шаге. Обратите внимание на то, что он присваивает параметру `forceConsent` значение `false`. Дополнительные сведения см. в описании следующего шага.</span><span class="sxs-lookup"><span data-stu-id="a8328-p122">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }
    ```

1. <span data-ttu-id="a8328-p123">Под методом `getOneDriveFiles` добавьте приведенный ниже код. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p123">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="a8328-215">[getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) — это новый API в Office.js, позволяющий надстройке запрашивать у ведущего приложения Office (Excel, PowerPoint, Word и т. д.) маркер доступа к надстройке (для пользователя, выполнившего вход в Office).</span><span class="sxs-lookup"><span data-stu-id="a8328-215">The [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office).</span></span> <span data-ttu-id="a8328-216">Ведущее приложение Office, в свою очередь, запрашивает маркер у конечной точки Azure AD версии 2.0.</span><span class="sxs-lookup"><span data-stu-id="a8328-216">The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token.</span></span> <span data-ttu-id="a8328-217">Так как вы предварительно авторизовали ведущее приложение Office для надстройки во время ее регистрации, Azure AD отправит токен.</span><span class="sxs-lookup"><span data-stu-id="a8328-217">Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="a8328-218">Если вход в Office не выполнен, ведущее приложение Office предложит пользователю войти.</span><span class="sxs-lookup"><span data-stu-id="a8328-218">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="a8328-p125">Параметр настроек задает для `forceConsent` значение `false`, поэтому пользователю не будет предлагаться разрешить ведущему приложению Office доступ к надстройке при каждом ее использовании. При первом запуске надстройки вызов `getAccessTokenAsync` не будет выполнен, но логика обработки ошибок, которую вы добавите на следующем этапе, автоматически выполнит повторный вызов, при этом параметру `forceConsent` будет задано значение `true`, и пользователю будет предложено согласиться. Такая процедура выполняется только в первый раз.</span><span class="sxs-lookup"><span data-stu-id="a8328-p125">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in. The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="a8328-221">Вы создадите метод `handleClientSideErrors` позже.</span><span class="sxs-lookup"><span data-stu-id="a8328-221">You will create the `handleClientSideErrors` method in a later step.</span></span>

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

1. <span data-ttu-id="a8328-p126">Замените строку TODO1 на приведенные ниже строки. Метод `getData` и серверный маршрут /api/values создаются позже. Для конечной точки используется относительный URL-адрес, так как она должна размещаться на том же домене, что и надстройка.</span><span class="sxs-lookup"><span data-stu-id="a8328-p126">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="a8328-p127">Под методом `getOneDriveFiles` добавьте приведенный ниже код. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p127">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="a8328-p128">Этот метод вызывает указанную конечную точку веб-API и передает ей тот же маркер доступа, который ведущее приложение Office использовало для доступа к надстройке. На стороне сервера этот маркер доступа будет использоваться в потоке "от имени" для получения маркера доступа к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a8328-p128">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="a8328-229">Вы создадите метод `handleServerSideErrors` позже.</span><span class="sxs-lookup"><span data-stu-id="a8328-229">You will create the `handleServerSideErrors` method in a later step.</span></span>

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

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="a8328-230">Создание методов обработки ошибок</span><span class="sxs-lookup"><span data-stu-id="a8328-230">Create the error-handling methods</span></span>

1. <span data-ttu-id="a8328-p129">Под методом `getData` добавьте приведенный ниже метод. Этот метод будет обрабатывать ошибки в клиенте надстройки, когда ведущее приложение Office не сможет получить маркер доступа к веб-службе надстройки. Сообщения о таких ошибках содержат код ошибки, поэтому данный метод различает их с помощью оператора `switch`.</span><span class="sxs-lookup"><span data-stu-id="a8328-p129">Below the `getData` method, add the following method. This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service. These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {

            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor.

            // TODO3: Handle the case where the user's sign-in or consent was aborted.

            // TODO4: Handle the case where the user is logged in with an account that is neither work or school,
            //        nor Microsoft Account.

            // TODO5: Handle the case where the Office host has not been authorized to the add-in's web service or
            //        the user has not granted the service permission to their `profile`.

            // TODO6: Handle an unspecified error from the Office host.

            // TODO7: Handle the case where the Office host cannot get an access token to the add-ins
            //        web service/application.

            // TODO8: Handle the case where the user triggered an operation that calls `getAccessTokenAsync`
            //        before a previous call of it completed.

            // TODO9: Handle the case where the add-in does not support forcing consent.

            // TODO10: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="a8328-p130">Замените `TODO2` приведенным ниже кодом. Ошибка 13001 возникает, если пользователь не выполнил вход или без отклика отменил запрос на предоставление 2-го фактора проверки подлинности. В обоих случаях код повторно выполняет метод `getDataWithToken` и задает параметр для принудительного запрашивания входа.</span><span class="sxs-lookup"><span data-stu-id="a8328-p130">Replace `TODO2` with the following code. Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor. In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="a8328-p131">Замените `TODO3` приведенным ниже кодом. Ошибка 13002 возникает, когда вход или предоставление разрешений прерывается. Попросите пользователя повторить попытку, но не более одного раза.</span><span class="sxs-lookup"><span data-stu-id="a8328-p131">Replace `TODO3` with the following code. Error 13002 occurs when user's sign-in or consent was aborted. Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }
        break;
    ```

1. <span data-ttu-id="a8328-240">Замените `TODO4` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-240">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="a8328-241">Ошибка 13003 возникает, когда пользователь входит под учетной записью, отличной от рабочей, учебной или личной учетной записи Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="a8328-241">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft account.</span></span> <span data-ttu-id="a8328-242">Попросите пользователя выйти, а затем войти с помощью учетной записи поддерживаемого типа.</span><span class="sxs-lookup"><span data-stu-id="a8328-242">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003:
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;
    ```

    > [!NOTE]
    > <span data-ttu-id="a8328-243">Ошибка 13004 не обрабатывается при использовании этого метода, так как она должна возникать только на стадии разработки.</span><span class="sxs-lookup"><span data-stu-id="a8328-243">Error 13004 is not handled in this method because it should only occur in development.</span></span> <span data-ttu-id="a8328-244">Ее невозможно исправить с помощью кода среды выполнения, поэтому нет смысла сообщать о ней пользователю.</span><span class="sxs-lookup"><span data-stu-id="a8328-244">It cannot be fixed by runtime code and there would be no point in reporting it to an end user.</span></span>

1. <span data-ttu-id="a8328-245">Замените `TODO5` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-245">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="a8328-246">Ошибка 13005 возникает, когда Office не имеет разрешение на использование надстройки веб-службы, либо пользователь не предоставил разрешение на использование службы для `profile`.</span><span class="sxs-lookup"><span data-stu-id="a8328-246">Error 13005 occurs when Office has not been authorized to the add-in's web service or the user has not granted the service permission to their `profile`.</span></span>

    ```javascript
    case 13005:
        getDataWithToken({ forceConsent: true });
        break;
    ```

1. <span data-ttu-id="a8328-p135">Замените `TODO6` приведенным ниже кодом. Ошибка 13006 возникает, если происходит неопределенная ошибка ведущего приложения Office, которая может свидетельствовать о его нестабильном состоянии. Попросите пользователя перезапустить Office.</span><span class="sxs-lookup"><span data-stu-id="a8328-p135">Replace `TODO6` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;
    ```

1. <span data-ttu-id="a8328-p136">Замените `TODO7` приведенным ниже кодом. Ошибка 13007 возникает, когда нарушается взаимодействие ведущего приложения Office с AAD, из-за чего это приложение не может получить маркер доступа к веб-службе/приложению надстройки. Это может быть из-за временного сбоя сети. Попросите пользователя повторить попытку позже.</span><span class="sxs-lookup"><span data-stu-id="a8328-p136">Replace `TODO7` with the following code. Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application. This may be a temporary network issue. Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;
    ```

1. <span data-ttu-id="a8328-254">Замените `TODO8` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-254">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="a8328-255">Ошибка 13008 возникает, когда пользователь запускает операцию, которая вызывает `getAccessTokenAsync`, до завершения предыдущего вызова.</span><span class="sxs-lookup"><span data-stu-id="a8328-255">Error 13008 occurs when the user triggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```

1. <span data-ttu-id="a8328-p138">Замените `TODO9` указанным ниже кодом. Ошибка 13009 возникает, если надстройка не поддерживает принудительное запрашивание разрешения, но выполняется вызов `getAccessTokenAsync` с установкой для параметра `forceConsent` значения `true`. Обычно в таком случае код должен автоматически повторно запустить метод `getAccessTokenAsync` с параметром, имеющим значение `false`. Но в некоторых случаях вызов метода с установкой для параметра `forceConsent` значения `true` сам по себе является автоматическим откликом на ошибку вызова метода с установкой для параметра значения `false`. В этом случае код должен не повторять попытку, а предложить пользователю выйти и войти заново.</span><span class="sxs-lookup"><span data-stu-id="a8328-p138">Replace `TODO9` with the following code. Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`. In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`. However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`. In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```

1. <span data-ttu-id="a8328-261">Замените `TODO10` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-261">Replace `TODO10` with the following code.</span></span>

    ```javascript
    default:
        logError(result);
        break;
    ```  


1. <span data-ttu-id="a8328-p139">Под методом `handleClientSideErrors` добавьте приведенный ниже метод. Этот метод обрабатывает ошибки в веб-службе надстройки при неправильном выполнении потока "от имени" или получении данных от Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a8328-p139">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {

        // TODO11: Parse the JSON response.

        // TODO12: Handle the case where AAD asks for an additional form of authentication.

        // TODO13: Handle missing consent and scope (permission) related issues.

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. <span data-ttu-id="a8328-264">Замените `TODO11` указанным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-264">Replace `TODO11` with the following code.</span></span> <span data-ttu-id="a8328-265">Обратите внимание, что для большинства ошибок `4xx`, которые веб-служба будет передавать клиентской части надстройки, в ответе будет свойство **ExceptionMessage**, содержащее номер ошибки AADSTS и другие данные.</span><span class="sxs-lookup"><span data-stu-id="a8328-265">Note that for most of the `4xx` errors that the add-in's web service will pass to the add-in's client-side, there will be an **ExceptionMessage** property in the response that contains the AADSTS (Azure Active Directory Secure Token Service) error number as well as other data.</span></span> <span data-ttu-id="a8328-266">Однако, когда AAD отправляет веб-службе надстройки запрос дополнительной проверки подлинности, этот запрос содержит специальное свойство **Claims** с кодом необходимой дополнительной проверки.</span><span class="sxs-lookup"><span data-stu-id="a8328-266">However, when AAD sends a message to the add-in's web service asking for an additional authentication factor, the message contains a special **Claims** property that specifies (with a code number) what additional factor is needed.</span></span> <span data-ttu-id="a8328-267">API ASP.NET, которые создают и отправляют HTTP-ответы клиентам, не знают об этом свойстве **Claims**, поэтому не включают его в ответ.</span><span class="sxs-lookup"><span data-stu-id="a8328-267">The ASP.NET APIs that create and send HTTP Responses to clients do not know about this **Claims** property, so they do not include it in the Response object.</span></span> <span data-ttu-id="a8328-268">Серверный код, который вы создадите позже, будет вручную добавлять значение **Claims** в ответ, чтобы решить эту проблему.</span><span class="sxs-lookup"><span data-stu-id="a8328-268">Server-side code that you will create in a later step will cope with this by manually adding the **Claims** value to the Response object.</span></span> <span data-ttu-id="a8328-269">Это значение будет находиться в свойстве **Message**, поэтому код также должен анализировать это свойство.</span><span class="sxs-lookup"><span data-stu-id="a8328-269">This value will be in the **Message** property, so the code needs to parse out that property as well.</span></span>

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. <span data-ttu-id="a8328-270">Замените `TODO12` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-270">Replace `TODO12` with the following code.</span></span> <span data-ttu-id="a8328-271">Что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-271">Note about this code:</span></span>

    * <span data-ttu-id="a8328-272">Ошибка 50076 возникает, когда Microsoft Graph требует дополнительной проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="a8328-272">Error 50076 occurs when Microsoft Graph requires an additional form of authentication.</span></span>
    * <span data-ttu-id="a8328-p142">Основное приложение Office должно получить новый маркер со значением **Claims** в качестве параметра `authChallenge`. В результате AAD предложит пользователю пройти все необходимые проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="a8328-p142">The Office host should get a new token with the **Claims** value as the `authChallenge` option. This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }
    ```

1. <span data-ttu-id="a8328-275">Замените `TODO13` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-275">Replace `TODO13` with the following code.</span></span> <span data-ttu-id="a8328-276">Вы замените три элемента `TODO` в этом коде с использованием *внутреннего* условного блока на следующих нескольких этапах.</span><span class="sxs-lookup"><span data-stu-id="a8328-276">You will replace the three `TODO`s in this code with an *inner* conditional block in the next few steps.</span></span>

    ```javascript
    else if (exceptionMessage) {

        // TODO13A: Handle the case where consent has not been granted, or has been revoked.

        // TODO13B: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO13C: Handle the case where the token that the add-in's client-side sends to it's
        //          server-side is not valid because it is missing `access_as_user` scope (permission).
    }
  
    ```


1. <span data-ttu-id="a8328-277">Замените `TODO13A` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-277">Replace `TODO13A` with the following code.</span></span> <span data-ttu-id="a8328-278">(Это создает первую часть *внутреннего* условного блока). Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-278">(This creates the first part of an *inner* conditional block.) Note about this code:</span></span>

    * <span data-ttu-id="a8328-279">Ошибка 65001 означает, что доступ к Microsoft Graph не был предоставлен (или был отозван) для одного или нескольких разрешений.</span><span class="sxs-lookup"><span data-stu-id="a8328-279">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span>
    * <span data-ttu-id="a8328-280">Надстройка должна получить новый маркер с параметром `forceConsent`, имеющим значение `true`.</span><span class="sxs-lookup"><span data-stu-id="a8328-280">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
       getDataWithToken({ forceConsent: true });
    }
    ```

1. <span data-ttu-id="a8328-p145">Замените `TODO13B` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p145">Replace `TODO13B` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="a8328-p146">Ошибка 70011 имеет несколько значений. Главное для этой надстройки — запрашивание недопустимого разрешения, поэтому код проверяет наличие полного описания ошибки, а не только номера.</span><span class="sxs-lookup"><span data-stu-id="a8328-p146">Error 70011 has multiple meanings. The one that matters to this add-in is when it means that an invalid scope (permission) has been requested, so the code checks for the full error description, not just the number.</span></span>
    * <span data-ttu-id="a8328-285">Надстройка должна сообщить об ошибке.</span><span class="sxs-lookup"><span data-stu-id="a8328-285">The add-in should report the error.</span></span>

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. <span data-ttu-id="a8328-p147">Замените `TODO13C` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p147">Replace `TODO13C` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="a8328-288">Серверный код, который вы создадите позже, отправит сообщение `Missing access_as_user`, если разрешения `access_as_user` не будет в маркере доступа, который клиент надстройки отправит в AAD для использования в потоке "от имени".</span><span class="sxs-lookup"><span data-stu-id="a8328-288">Server-side code that you create in a later step will send the message `Missing access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="a8328-289">Надстройка должна сообщить об ошибке.</span><span class="sxs-lookup"><span data-stu-id="a8328-289">The add-in should report the error.</span></span>

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. <span data-ttu-id="a8328-290">Замените `TODO14` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-290">Replace `TODO14` with the following code.</span></span> <span data-ttu-id="a8328-291">(Этот код является частью *внешнего* условного блока и должен следовать сразу же после закрывающихся скобок структуры, которая начинается с `else if (exceptionMessage) {` и на одном уровне отступа). Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-291">(This is part of the *outer* conditional block and should be immediately after the close bracket of the structure that begins with `else if (exceptionMessage) {` and at the same level of indentation.) Note about this code:</span></span>

    * <span data-ttu-id="a8328-292">Библиотека идентификации, которую вы будете использовать в серверном коде (MSAL), должна предотвращать отправку в Microsoft Graph устаревших и недействительных маркеров; но если это все-таки произойдет, Microsoft Graph вернет веб-службе надстройки ошибку с кодом `InvalidAuthenticationToken`.</span><span class="sxs-lookup"><span data-stu-id="a8328-292">The identity library that you will be using in the server-side code (Microsoft Authentication Library - MSAL) should ensure that no expired or invalid token is sent to Microsoft Graph; but if it does happen, the error that is returned to the add-in's web service from Microsoft Graph has the code `InvalidAuthenticationToken`.</span></span> <span data-ttu-id="a8328-293">Серверный код, который вы создадите позже, передаст это сообщение клиенту надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8328-293">Server-side code you will create in a later step will relay this message to the add-in's client.</span></span>
    * <span data-ttu-id="a8328-294">В этом случае надстройка должна начать заново весь процесс проверки подлинности, сбросив счетчик и переменные флага, а затем повторно вызвать метод обработчика кнопок.</span><span class="sxs-lookup"><span data-stu-id="a8328-294">In this case, the add-in should start the entire authentication process over by resetting the counter and flag variables, and then re-calling the button handler method.</span></span>

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }
    ```

1. <span data-ttu-id="a8328-295">Замените `TODO15` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="a8328-295">Replace `TODO15` with the following code.</span></span>

    ```javascript
    else {
        logError(result);
    }
    ```

1. <span data-ttu-id="a8328-296">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="a8328-296">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="a8328-297">Код на стороне сервера</span><span class="sxs-lookup"><span data-stu-id="a8328-297">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="a8328-298">Настройка ПО промежуточного слоя OWIN</span><span class="sxs-lookup"><span data-stu-id="a8328-298">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="a8328-299">Откройте файл Startup.cs в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="a8328-299">Open the Startup.cs file in the root of the project.</span></span>

1. <span data-ttu-id="a8328-p150">Добавьте ключевое слово `partial` в объявление класса Startup, если его там еще нет. Оно должно выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="a8328-p150">Add the keyword `partial` to the declaration of the Startup class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="a8328-p151">Добавьте приведенную ниже строку в текст метода `Configuration`. Метод `ConfigureAuth` создается позже.</span><span class="sxs-lookup"><span data-stu-id="a8328-p151">Add the following line to the body of the `Configuration` method. You create the `ConfigureAuth` method in a later step.</span></span>

    `ConfigureAuth(app);`

1. <span data-ttu-id="a8328-304">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="a8328-304">Save and close the file.</span></span>

1. <span data-ttu-id="a8328-305">Щелкните правой кнопкой мыши папку **App_Start** и выберите **Добавить > Класс**.</span><span class="sxs-lookup"><span data-stu-id="a8328-305">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="a8328-306">В диалоговом окне **Добавить новый элемент** введите имя файла **Startup.Auth.cs** и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="a8328-306">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="a8328-307">Сократите имя пространства имен в новом файле до `Office_Add_in_ASPNET_SSO_WebAPI`.</span><span class="sxs-lookup"><span data-stu-id="a8328-307">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="a8328-308">Убедитесь, что в начале файла есть все приведенные ниже операторы `using`.</span><span class="sxs-lookup"><span data-stu-id="a8328-308">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="a8328-p152">Добавьте ключевое слово `partial` в объявление класса `Startup`, если его там еще нет. Оно должно выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="a8328-p152">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="a8328-p153">Добавьте приведенный ниже метод в класс `Startup`. Этот метод указывает, как ПО промежуточного слоя OWIN будет проверять маркеры доступа, передаваемые ему из метода `getData` в файле Home.js на стороне клиента. Процесс вызывается при каждом вызове конечной точки веб-API, содержащей атрибут `[Authorize]`.</span><span class="sxs-lookup"><span data-stu-id="a8328-p153">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. <span data-ttu-id="a8328-p154">Замените TODO3 приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p154">Replace the TODO3 with the following. Note about this code:</span></span>

    * <span data-ttu-id="a8328-316">Код сообщает OWIN о необходимости убедиться, что аудитория и поставщик маркера, указанные в маркере доступа из ведущего приложения Office (который передается путем вызова метода `getData` на стороне клиента), должны совпадать со значениями, указанными в файле web.config.</span><span class="sxs-lookup"><span data-stu-id="a8328-316">The code instructs OWIN to ensure that the audience and token issuer specified in the access token that comes from the Office host (and is passed on by the client-side call of `getData`) must match the values specified in the web.config.</span></span>
    * <span data-ttu-id="a8328-p155">Если задать для свойства `SaveSigninToken` значение `true`, OWIN сохранит необработанный маркер из ведущего приложения Office. Он необходим надстройке, чтобы получить маркер доступа к Microsoft Graph в потоке "от имени".</span><span class="sxs-lookup"><span data-stu-id="a8328-p155">Setting `SaveSigninToken` to `true` causes OWIN to save the raw token from the Office host. The add-in needs it to obtain an access token to Microsoft Graph with the “on behalf of” flow.</span></span>
    * <span data-ttu-id="a8328-p156">ПО промежуточного слоя OWIN не проверяет разрешения. Разрешения маркера доступа, которые должны включать `access_as_user`, проверяются в контроллере.</span><span class="sxs-lookup"><span data-stu-id="a8328-p156">Scopes are not validated by the OWIN middleware. The scopes of the access token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. <span data-ttu-id="a8328-p157">Замените TODO4 приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p157">Replace TODO4 with the following. Note about this code:</span></span>

    * <span data-ttu-id="a8328-323">Метод `UseOAuthBearerAuthentication` вызывается вместо более распространенного метода `UseWindowsAzureActiveDirectoryBearerAuthentication`, так как последний несовместим с конечной точкой Azure AD версии 2.</span><span class="sxs-lookup"><span data-stu-id="a8328-323">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="a8328-324">ПО промежуточного слоя OWIN использует URL-адрес обнаружения, передаваемый методу, чтобы получить ключ, необходимый для проверки подписи в маркере доступа, полученном из ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="a8328-324">The discovery URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the access token received from the Office host.</span></span>

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. <span data-ttu-id="a8328-325">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="a8328-325">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="a8328-326">Создание контроллера /api/values</span><span class="sxs-lookup"><span data-stu-id="a8328-326">Create the /api/values controller</span></span>

1. <span data-ttu-id="a8328-327">Откройте файл **Controllers\ValueController.cs**.</span><span class="sxs-lookup"><span data-stu-id="a8328-327">Open the file **Controllers\ValueController.cs**.</span></span>

1. <span data-ttu-id="a8328-328">Убедитесь, что в начале файла есть приведенные ниже инструкции с `using`.</span><span class="sxs-lookup"><span data-stu-id="a8328-328">Ensure that the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

1. <span data-ttu-id="a8328-p158">Над строкой с объявлением `ValuesController` добавьте атрибут `[Authorize]`. Это гарантирует, что надстройка будет выполнять процесс авторизации, настроенный в последней процедуре, при каждом вызове метода контроллера. Вызывать методы контроллера можно только при наличии действительного маркера доступа к надстройке.</span><span class="sxs-lookup"><span data-stu-id="a8328-p158">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a8328-p159">Производственная служба веб-API на основе ASP.NET MVC должна иметь специальную логику для потока "от имени" в одном или нескольких пользовательских классах **FilterAttribute**. В этом примере логика помещается в главный контроллер, чтобы можно было легко проследить весь поток авторизации и логику получения данных. Такая же модель используется в примерах авторизации в разделе [Azure Samples](https://github.com/Azure-Samples/).</span><span class="sxs-lookup"><span data-stu-id="a8328-p159">A production ASP.NET MVC Web API service should have custom logic for the on-behalf-of flow in one or more custom **FilterAttribute** classes. This educational sample puts the logic in the main controller so that the entire flow of the authorization and data fetching logic can be easily followed. This also makes the sample consistent with the pattern of authorization samples in [Azure Samples](https://github.com/Azure-Samples/).</span></span>

1. <span data-ttu-id="a8328-p160">Добавьте приведенный ниже метод в `ValuesController`. Обратите внимание, что возвращаемое значение — `Task<HttpResponseMessage>`, а не `Task<IEnumerable<string>>`, которое чаще используется для метода `GET api/values`. Это побочный эффект нахождения пользовательской логики авторизации в контроллере: при возникновении некоторых ошибок веб-служба должна отправлять HTTP-ответ клиенту надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8328-p160">Add the following method to the `ValuesController`. Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method. This is a side effect of that fact that our custom authorization logic will be in the controller: some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span>

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

1. <span data-ttu-id="a8328-338">Замените `TODO1` приведенным ниже кодом, чтобы убедиться, что в маркере указано разрешение `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="a8328-338">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span>

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > <span data-ttu-id="a8328-339">Для авторизации API, который отвечает за поток выполнения от имени другого субъекта, в случае надстроек Office используйте только область `access_as_user`. Для других API в службе должны быть предусмотрены отдельные требования, касающиеся областей.</span><span class="sxs-lookup"><span data-stu-id="a8328-339">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements.</span></span> <span data-ttu-id="a8328-340">Это ограничивает доступ, предоставляемый с использованием маркеров, которые получает Office.</span><span class="sxs-lookup"><span data-stu-id="a8328-340">This limits what can be accessed with the tokens that Office acquires.</span></span>

1. <span data-ttu-id="a8328-p162">Замените `TODO2` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p162">Replace `TODO2` with the following code. Note about this code:</span></span>
    * <span data-ttu-id="a8328-343">Код преобразует необработанный маркер доступа, полученный от ведущего приложения Office, в объект `UserAssertion`, который будет передан другому методу.</span><span class="sxs-lookup"><span data-stu-id="a8328-343">It turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.</span></span>
    * <span data-ttu-id="a8328-p163">Надстройка больше не выступает в роли ресурса (или аудитории), доступ к которому необходим ведущему приложению Office и пользователю. Теперь она сама является клиентом, которому необходим доступ к Microsoft Graph. `ConfidentialClientApplication` — это объект "контекста клиента" MSAL.</span><span class="sxs-lookup"><span data-stu-id="a8328-p163">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="a8328-p164">Третий параметр конструктора `ConfidentialClientApplication` — URL-адрес перенаправления. На самом деле он не используется в потоке "от имени", но все равно рекомендуется указывать правильный URL-адрес. С помощью четвертого и пятого параметров можно определить постоянное хранилище, которое позволяет повторно использовать действительные маркеры в разных сеансах с надстройкой. В этом примере не реализуется постоянное хранилище.</span><span class="sxs-lookup"><span data-stu-id="a8328-p164">The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the “on behalf of” flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.</span></span>
    * <span data-ttu-id="a8328-p165">Для работы библиотеки MSAL требуются области `openid` и `offline_access`, но если код их избыточно запрашивает, возникает ошибка. Кроме того, ошибка возникнет, если код запросит `profile` (фактически используется только при получении ведущим приложением Office токена для веб-приложения надстройки). Поэтому явным образом запрашивается только `Files.Read.All`.</span><span class="sxs-lookup"><span data-stu-id="a8328-p165">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them. It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application. So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

1. <span data-ttu-id="a8328-p166">Замените `TODO3` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p166">Replace `TODO3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="a8328-p167">Для начала метод `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` проверит кэш MSAL, который находится в памяти, на наличие подходящего маркера доступа. Только в случае его отсутствия запускается поток "от имени" с конечной точкой Azure AD версии 2.</span><span class="sxs-lookup"><span data-stu-id="a8328-p167">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="a8328-357">Если ресурс Microsoft Graph требует многофакторной проверки подлинности, а пользователь еще не предоставил соответствующие данные, AAD вызовет исключение, содержащее свойство Claims.</span><span class="sxs-lookup"><span data-stu-id="a8328-357">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will throw an exception containing a Claims property.</span></span>
    * <span data-ttu-id="a8328-p168">Значение свойства Claims необходимо передать клиенту, который передаст его ведущему приложению Office. Последнее добавит его в запрос на получение нового токена. AAD предложит пользователю пройти все необходимые проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="a8328-p168">The Claims property value must be passed to the client which will pass it to the Office host, which will then include it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="a8328-360">Любые исключения, отличные от типа `MsalServiceException`, не перехватываются преднамеренно, поэтому будут переданы клиенту в виде сообщений `500 Server Error`.</span><span class="sxs-lookup"><span data-stu-id="a8328-360">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

1. <span data-ttu-id="a8328-p169">Замените `TODO3a` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p169">Replace `TODO3a` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="a8328-p170">Если ресурс Microsoft Graph требует многофакторной проверки подлинности, а пользователь еще не предоставил соответствующие данные, AAD вернет состояние "400 Bad Request" с ошибкой AADSTS50076 и свойство **Claims**. MSAL выдает исключение **MsalUiRequiredException** (наследуется от исключения **MsalServiceException**) с этой информацией.</span><span class="sxs-lookup"><span data-stu-id="a8328-p170">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will return "400 Bad Request" with error AADSTS50076 and a **Claims** property. MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span> 
    * <span data-ttu-id="a8328-p171">Значение свойства **Claims** необходимо передать клиенту, который передаст его ведущему приложению Office. Последнее добавит его в запрос на получение нового токена. AAD предложит пользователю пройти все необходимые проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="a8328-p171">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="a8328-p172">API, которые создают HTTP-ответы из исключений, не знают о свойстве **Claims**, поэтому не включают его в ответ. Нам нужно создать сообщение с ним вручную. Однако настраиваемое свойство **Message** блокирует создание свойства **ExceptionMessage**, поэтому единственный способ передать идентификатор ошибки `AADSTS50076` клиенту — добавить его в настраиваемое свойство **Message**. Код JavaScript в клиенте должен будет определить, какое свойство содержится в ответе (**Message** или **ExceptionMessage**).</span><span class="sxs-lookup"><span data-stu-id="a8328-p172">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="a8328-371">Сообщение создается в формате JSON, чтобы клиентский код JavaScript мог проанализировать его с помощью известных методов объекта `JSON`.</span><span class="sxs-lookup"><span data-stu-id="a8328-371">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known `JSON` object methods.</span></span>
    * <span data-ttu-id="a8328-p173">Вы создадите метод `SendErrorToClient` позже. Его второй параметр — объект **Exception**. В этом случае код передает `null`, потому что включение объекта **Exception** блокирует включение свойства **Message** в создаваемый HTTP-ответ.</span><span class="sxs-lookup"><span data-stu-id="a8328-p173">You will create the `SendErrorToClient` method in a later step. It's second parameter is an **Exception** object. In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. <span data-ttu-id="a8328-p174">Замените `TODO3b` и `TODO3c` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p174">Replace `TODO3b` and `TODO3c` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="a8328-p175">Если вызов AAD содержал по крайней мере одно разрешение, которое не предоставил ни пользователь, ни администратор клиента (или оно было отозвано), AAD вернет состояние "400 Bad Request" с ошибкой `AADSTS65001`. MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию. Клиент должен вызвать метод `getAccessTokenAsync` повторно, используя параметр `{ forceConsent: true }`.</span><span class="sxs-lookup"><span data-stu-id="a8328-p175">If the call to AAD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked). AAD will return "400 Bad Request" with error `AADSTS65001`. MSAL throws a **MsalUiRequiredException** with this information. The client should re-call `getAccessTokenAsync` with the option `{ forceConsent: true }`.</span></span>
    *  <span data-ttu-id="a8328-p176">Если вызов AAD содержал по крайней мере одно нераспознанное разрешение, AAD вернет состояние "400 Bad Request" с ошибкой `AADSTS70011`. MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию. Клиент должен сообщить об этом пользователю.</span><span class="sxs-lookup"><span data-stu-id="a8328-p176">If the call to AAD contained at least one scope that AAD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`. MSAL throws a **MsalUiRequiredException** with this information. The client should inform the user.</span></span>
    *  <span data-ttu-id="a8328-384">Полное описание включается, так как ошибка 70011 возвращается и в других случаях, и ее следует обрабатывать в этой надстройке, только когда она означает запрос недопустимого разрешения.</span><span class="sxs-lookup"><span data-stu-id="a8328-384">The entire description is included because 70011 is returned in other conditions and we it should only be handled in this add-in when it means that there is an invalid scope.</span></span>
    *  <span data-ttu-id="a8328-p177">Объект **MsalUiRequiredException** передается методу `SendErrorToClient`. Это гарантирует, что свойство **ExceptionMessage**, содержащее информацию об ошибке, будет включено в HTTP-отклик.</span><span class="sxs-lookup"><span data-stu-id="a8328-p177">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>
    *  <span data-ttu-id="a8328-387">Сообщения нет, поэтому в качестве третьего параметра передается `null`.</span><span class="sxs-lookup"><span data-stu-id="a8328-387">There is no custom message, so `null` is passed for the third parameter.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. <span data-ttu-id="a8328-p178">Замените `TODO3d` приведенным ниже кодом. Обратите внимание, что код повторно выдает исключение, а не передает его в собственном HTTP-ответе с состоянием **HttpStatusCode.Forbidden** (401). В результате ASP.NET отправляет собственный HTTP-ответ с состоянием "500 Server Error".</span><span class="sxs-lookup"><span data-stu-id="a8328-p178">Replace `TODO3d` with the following code. Note that the code rethrows the exception instead of relaying it in a custom HTTP Response with **HttpStatusCode.Forbidden** (401). The effect of this is that the ASP.NET will send its own HTTP Response with status "500 Server Error".</span></span>

    ```csharp
    else
    {
        throw e;
    }  
    ```

1. <span data-ttu-id="a8328-p179">Замените `TODO4` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p179">Replace `TODO4` with the following. Note about this code:</span></span>

    * <span data-ttu-id="a8328-p180">Классы `GraphApiHelper` и `ODataHelper` определяются в файлах из папки **Helpers**. Класс `OneDriveItem` определяется в файле из папки **Models**. В этой статье не представлено подробное описание этих классов, так как оно не имеет отношения к авторизации и единому входу.</span><span class="sxs-lookup"><span data-stu-id="a8328-p180">The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.</span></span>
    * <span data-ttu-id="a8328-396">Производительность будет выше, если запрашивать у Microsoft Graph только действительно необходимые данные, поэтому в коде заданы параметры `$select` и `$top`. Первый из них показывает, что нужно только свойство name, второй — что требуются только первые три названия папок или файлов.</span><span class="sxs-lookup"><span data-stu-id="a8328-396">Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a `$select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first three folder or file names.</span></span>
    * <span data-ttu-id="a8328-p181">Если отправленный в Microsoft Graph токен недействителен, Microsoft Graph возвращает ошибку "401 Unauthorized" с кодом "InvalidAuthenticationToken". ASP.NET затем выдает исключение **RuntimeBinderException**. Это также происходит, когда срок действия токена истек, хотя MSAL должна предотвращать отправку таких токенов.</span><span class="sxs-lookup"><span data-stu-id="a8328-p181">If the token sent to Microsoft Graph is invalid, Microsoft Graph sends a "401 Unauthorized" error with the code "InvalidAuthenticationToken". ASP.NET then throws a **RuntimeBinderException**. This is also what happens when the token is expired, although MSAL should prevent that from ever happening.</span></span> 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);
    }
    ```

1. <span data-ttu-id="a8328-p182">Замените `TODO5` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p182">Replace `TODO5` with the following. Note about this code:</span></span>

    * <span data-ttu-id="a8328-p183">Хотя приведенный выше код запрашивает только свойство *name* элементов OneDrive, Microsoft Graph всегда включает свойство *eTag* для элементов OneDrive. Чтобы сократить количество полезных данных, отправляемых клиенту, приведенный ниже код преобразует результаты, оставляя только имена элементов.</span><span class="sxs-lookup"><span data-stu-id="a8328-p183">Although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.</span></span>
    * <span data-ttu-id="a8328-404">Список из трех файлов и папок OneDrive отправляется клиенту в виде HTTP-ответа "200 OK".</span><span class="sxs-lookup"><span data-stu-id="a8328-404">The list of three OneDrive files and folders is sent to the client as a "200 OK" HTTP Response.</span></span>

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames);
    return response;
    ```

1. <span data-ttu-id="a8328-p184">Добавьте приведенный ниже метод под методом Get. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8328-p184">Below the Get method, add the following method. About this code note:</span></span>  

    * <span data-ttu-id="a8328-407">Метод передает клиенту информацию об исключении на стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="a8328-407">The method relays to the client information about a server-side exception.</span></span>
    * <span data-ttu-id="a8328-408">Если методу будет передано исходное исключение, конструктор HttpError включит информацию из исключения в свойство **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="a8328-408">If the original exception is passed to the method, then the HttpError constructor will include information from the exception object in an **ExceptionMessage** property.</span></span>  
    * <span data-ttu-id="a8328-409">Если в виде исключения будет передано значение `null`, конструктор HttpError включит параметр message в свойство **Message**. Свойства **ExceptionMessage** не будет.</span><span class="sxs-lookup"><span data-stu-id="a8328-409">If `null` is passed for the exception, then the HttpError constructor will include the message parameter in a **Message** property and there is no **ExceptionMessage** property.</span></span>

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }
    ```

## <a name="run-the-add-in"></a><span data-ttu-id="a8328-410">Запуск надстройки</span><span class="sxs-lookup"><span data-stu-id="a8328-410">Run the add-in</span></span>

1. <span data-ttu-id="a8328-411">Убедитесь в наличии нескольких файлов в OneDrive, чтобы можно было проверить результаты.</span><span class="sxs-lookup"><span data-stu-id="a8328-411">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="a8328-p185">В Visual Studio нажмите клавишу F5. Откроется PowerPoint, где на ленте **Главная** появится группа **SSO ASP.NET**.</span><span class="sxs-lookup"><span data-stu-id="a8328-p185">In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.</span></span>

1. <span data-ttu-id="a8328-414">Нажмите кнопку **Show Add-in** (Показать надстройку) в этой группе, чтобы увидеть пользовательский интерфейс надстройки в области задач.</span><span class="sxs-lookup"><span data-stu-id="a8328-414">Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane.</span></span>

1. <span data-ttu-id="a8328-p186">Нажмите кнопку **Get My Files from OneDrive** (Получить мои файлы из OneDrive). Если вы не вошли в Office, вам будет предложено войти.</span><span class="sxs-lookup"><span data-stu-id="a8328-p186">Press the button **Get My Files from OneDrive**. If you are not signed into Office, you'll be prompted to sign in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a8328-p187">Если ранее вы вошли в Office, используя другой идентификатор, и не закрыли некоторые из открытых тогда приложений Office, Office может не сменить идентификатор (даже если кажется, что это сделано для PowerPoint). Если это произойдет, возможен сбой при вызове Microsoft Graph или возврат данных для другого идентификатора. Чтобы избежать этого, *закройте все приложения Office*, прежде чем нажимать кнопку **Get My Files from OneDrive** (Получить мои файлы из OneDrive).</span><span class="sxs-lookup"><span data-stu-id="a8328-p187">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>

1. <span data-ttu-id="a8328-p188">После входа под кнопкой появится список файлов и папок из OneDrive. Это может занять более 15 секунд, особенно в первый раз.</span><span class="sxs-lookup"><span data-stu-id="a8328-p188">After you are signed in, a list of your files and folders on OneDrive will appear below the button. This may take over 15 seconds, especially the first time.</span></span>
