---
title: Создание такой надстройки Office на платформе ASP.NET, для которой используется единый вход
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: cdf039e66f0d61e656827ee3ab0ad5762cba430d
ms.sourcegitcommit: 8333ede51307513312d3078cb072f856f5bef8a2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/07/2018
ms.locfileid: "23876622"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="24359-102">Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="24359-102">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="24359-p101">Ваша надстройка может предоставлять пользователям доступ к нескольким приложениям, используя учетные данные, введенные при входе в Office. [Как включить единый вход в надстройке Office](sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="24359-p101">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="24359-105">Из этой статьи вы узнаете, как включить единый вход в надстройке, созданной с помощью ASP.NET, OWIN и MSAL для .NET.</span><span class="sxs-lookup"><span data-stu-id="24359-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET.</span></span>

> [!NOTE]
> <span data-ttu-id="24359-106">Сведения о создании надстройки, в которой используется единый вход, на основе Node.js см. в [этой статье](create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="24359-106">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="24359-107">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="24359-107">Prerequisites</span></span>

* <span data-ttu-id="24359-108">Последняя доступная версия Visual Studio 2017 Preview.</span><span class="sxs-lookup"><span data-stu-id="24359-108">The latest available version of Visual Studio 2017 Preview.</span></span>

* <span data-ttu-id="24359-p102">Office 2016 (версия 1708, сборка 8424.nnnn) или более поздняя (эту версию подписки на Office 365 иногда называют "нажми и работай"). Чтобы скачать эту версию, вам может потребоваться принять участие в программе предварительной оценки Office. Дополнительные сведения см. в статье [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="24359-p102">Office 2016, Version 1708, build 8424.nnnn or later (the Office 365 subscription version, sometimes called “Click to Run”). You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="24359-112">Настройка начального проекта</span><span class="sxs-lookup"><span data-stu-id="24359-112">Set up the starter project</span></span>

1. <span data-ttu-id="24359-113">Клонируйте или скачайте репозиторий [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span><span class="sxs-lookup"><span data-stu-id="24359-113">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

1. <span data-ttu-id="24359-p103">Перейдите в папку **Before** и откройте SLN-файл в Visual Studio. Это начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы.</span><span class="sxs-lookup"><span data-stu-id="24359-p103">Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.</span></span>

    > [!NOTE]
    > <span data-ttu-id="24359-p104">В том же репозитории есть готовая версия примера. Она идентична надстройке, которую вы создадите, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью. Чтобы использовать готовую версию, просто откройте файл `sln` и выполните действия, описанные в этой статье, пропустив разделы **Код на стороне клиента** и **Код на стороне сервера**.</span><span class="sxs-lookup"><span data-stu-id="24359-p104">There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures in this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the `sln` file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side.</span></span>

1. <span data-ttu-id="24359-p105">Открыв проект, выполните его сборку в Visual Studio. При этом будут установлены пакеты, указанные в файле packages.config. Это может занять от пары секунд до нескольких минут в зависимости от того, сколько пакетов хранится в локальном кэше пакетов на компьютере.</span><span class="sxs-lookup"><span data-stu-id="24359-p105">After the project opens, build it in Visual Studio, which will install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.</span></span>

    > [!NOTE]
    > <span data-ttu-id="24359-122">Вы увидите сообщение об ошибке, касающейся пространства имен Identity.</span><span class="sxs-lookup"><span data-stu-id="24359-122">You will get an error about the Identity namespace.</span></span> <span data-ttu-id="24359-123">Это побочный эффект проблемы с конфигурацией, которую вы устраните на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="24359-123">This is a side effect of a configuration issue that you will fix with the next step.</span></span> <span data-ttu-id="24359-124">Важно то, что пакеты устанавливаются.</span><span class="sxs-lookup"><span data-stu-id="24359-124">The important thing is that the packages are installed.</span></span>

1. <span data-ttu-id="24359-125">В настоящий момент версия библиотеки MSAL (Microsoft.Identity.Client), которая нужна для единого входа (версия `1.1.4-preview0002`), не включена в стандартный каталог NuGet, поэтому не указана в package.config. Ее нужно установить отдельно.</span><span class="sxs-lookup"><span data-stu-id="24359-125">Currently, the version of the MSAL library (Microsoft.Identity.Client) that you need for SSO (version `1.1.4-preview0002`) is not part of the standard nuget catalog, so it is not listed in the package.config, and it must be installed separately.</span></span> 

   > 1. <span data-ttu-id="24359-126">В меню **Сервис** выберите **Диспетчер пакетов NuGet** > **Консоль диспетчера пакетов**.</span><span class="sxs-lookup"><span data-stu-id="24359-126">On the **Tools** menu, navigate to **Nuget Package Manager** > **Package Manager Console**.</span></span> 

   > 2. <span data-ttu-id="24359-127">В консоли выполните указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="24359-127">At the console, run the following command.</span></span> <span data-ttu-id="24359-128">Выполнение может занять минуту или больше времени, даже при быстром подключении к Интернету.</span><span class="sxs-lookup"><span data-stu-id="24359-128">It may take a minute or more to complete even with a fast Internet connection.</span></span> <span data-ttu-id="24359-129">Когда все будет готово, в нижней части окна консоли отобразится такое сообщение: **"Microsoft.Identity.Client 1.1.4-preview0002" успешно установлено...**.</span><span class="sxs-lookup"><span data-stu-id="24359-129">When it finishes you should see **Successfully installed 'Microsoft.Identity.Client 1.1.1-alpha0393' ...** near the end of the output in the console.</span></span>

   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`

   > 3. <span data-ttu-id="24359-130">В **обозревателе решений** щелкните правой кнопкой мыши узел **Ссылки**.</span><span class="sxs-lookup"><span data-stu-id="24359-130">In **Solution Explorer**, right-click **References**.</span></span> <span data-ttu-id="24359-131">Убедитесь, что в него включена библиотека **Microsoft.Identity.Client**.</span><span class="sxs-lookup"><span data-stu-id="24359-131">Verify that **Microsoft.Identity.Client** is listed.</span></span> <span data-ttu-id="24359-132">Если ее нет (если она есть, но рядом с ней отображается значок предупреждения, удалите эту запись), добавьте ссылку в сборку с помощью мастера добавления ссылок Visual Studio, указав **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**</span><span class="sxs-lookup"><span data-stu-id="24359-132">If it is not or there is a warning icon on its entry, delete the entry and then use the Visual Studio Add Reference wizard to add a reference to the assembly at **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.1-alpha0393\lib\net45\Microsoft.Identity.Client.dll**</span></span>

1. <span data-ttu-id="24359-133">Еще раз выполните сборку проекта.</span><span class="sxs-lookup"><span data-stu-id="24359-133">Build the project a second time.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="24359-134">Регистрация надстройки в конечной точке Azure AD версии 2.0</span><span class="sxs-lookup"><span data-stu-id="24359-134">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="24359-135">Следующие инструкции содержат общую информацию, поэтому их можно использовать в нескольких местах.</span><span class="sxs-lookup"><span data-stu-id="24359-135">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="24359-136">Для этой статьи сделайте следующее:</span><span class="sxs-lookup"><span data-stu-id="24359-136">For this ariticle do the following:</span></span>
- <span data-ttu-id="24359-137">Замените заполнитель **$ADD-IN-NAME$** на `Office-Add-in-ASPNET-SSO`.</span><span class="sxs-lookup"><span data-stu-id="24359-137">Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-ASPNET-SSO`.</span></span>
- <span data-ttu-id="24359-138">Замените заполнитель **$FQDN-WITHOUT-PROTOCOL$** на `localhost:44355`.</span><span class="sxs-lookup"><span data-stu-id="24359-138">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:44355`.</span></span>
- <span data-ttu-id="24359-139">При указании разрешений в диалоговом окне **Выбрать Разрешения** установите флажки для следующих разрешений.</span><span class="sxs-lookup"><span data-stu-id="24359-139">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="24359-140">Для самой надстройки требуется только первое разрешение, а `offline_access` и `openid` требуются для библиотеки MSAL, используемой кодом на стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="24359-140">Only the first is really required by your add-in itself; but the MSAL library that the server-side code uses requires `offline_access` and `openid`.</span></span> <span data-ttu-id="24359-141">Разрешение `profile` необходимо, чтобы ведущее приложение Office получило токен для веб-приложения надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-141">The `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
    * <span data-ttu-id="24359-142">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="24359-142">Files.Read.All</span></span>
    * <span data-ttu-id="24359-143">offline_access</span><span class="sxs-lookup"><span data-stu-id="24359-143">offline_access</span></span>
    * <span data-ttu-id="24359-144">openid</span><span class="sxs-lookup"><span data-stu-id="24359-144">openid</span></span>
    * <span data-ttu-id="24359-145">profile</span><span class="sxs-lookup"><span data-stu-id="24359-145">profile</span></span>


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="24359-146">Предоставление надстройке согласия администратора</span><span class="sxs-lookup"><span data-stu-id="24359-146">Details are at: Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="24359-147">Конфигурация надстройки</span><span class="sxs-lookup"><span data-stu-id="24359-147">Configure the add-in</span></span>

1. <span data-ttu-id="24359-148">В следующей строке замените заполнитель "{tenant_ID}" на идентификатор клиента Office 365.</span><span class="sxs-lookup"><span data-stu-id="24359-148">In the following string, replace the placeholder “{tenant_ID}” with your Office 365 tenant ID.</span></span> <span data-ttu-id="24359-149">Для его получения используйте один из методов, описанных в статье [Как найти свой идентификатор клиента Office 365](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id).</span><span class="sxs-lookup"><span data-stu-id="24359-149">Use one of the methods in [Find your Office 365 tenant ID](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span>

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

2. <span data-ttu-id="24359-150">В Visual Studio откройте файл web.config. В разделе **appSettings** есть ключи, которым необходимо назначить значения.</span><span class="sxs-lookup"><span data-stu-id="24359-150">In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.</span></span>

3. <span data-ttu-id="24359-p112">Используйте строку, составленную на шаге 1, в качестве значения ключа ida:Issuer. Убедитесь, что в значении нет пробелов.</span><span class="sxs-lookup"><span data-stu-id="24359-p112">Use the string you constructed in step 1 as the value to the key named “ida:Issuer”. Be sure there are no blank spaces in the value.</span></span>

4. <span data-ttu-id="24359-153">Введите указанные ниже значения для соответствующих ключей.</span><span class="sxs-lookup"><span data-stu-id="24359-153">Assign the following values to the corresponding keys:</span></span>

    |<span data-ttu-id="24359-154">Ключ</span><span class="sxs-lookup"><span data-stu-id="24359-154">Key</span></span>|<span data-ttu-id="24359-155">Значение</span><span class="sxs-lookup"><span data-stu-id="24359-155">Value</span></span>|
    |:-----|:-----|
    |<span data-ttu-id="24359-156">ida:ClientID</span><span class="sxs-lookup"><span data-stu-id="24359-156">ida:ClientID</span></span>|<span data-ttu-id="24359-157">Идентификатор приложения, полученный во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-157">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="24359-158">ida:Audience</span><span class="sxs-lookup"><span data-stu-id="24359-158">ida:Audience</span></span>|<span data-ttu-id="24359-159">Идентификатор приложения, полученный во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-159">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="24359-160">ida:Password</span><span class="sxs-lookup"><span data-stu-id="24359-160">ida:Password</span></span>|<span data-ttu-id="24359-161">Пароль, который вы получили во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-161">TThe password you obtained when you registered the add-in.</span></span>|

   <span data-ttu-id="24359-p113">Ниже показан пример того, как должны выглядеть четыре измененные вами ключа. *Обратите внимание, что параметры ClientID и Audience имеют одинаковые значения*. Вы также можете использовать один ключ для обеих целей, но вашу разметку web.config будет проще повторно использовать, если вы разделите их, так как они не всегда будут одинаковыми. Кроме того, наличие отдельных ключей позволяет считать вашу надстройку и ресурсом OAuth, связанным с ведущим приложением Office, и клиентом OAuth, связанным с Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="24359-p113">The following is an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*. You can also use a single key for both purposes, but your web.config markup is more reusable if you keep them separate because they aren't always the same. Also, having separate keys reinforces the idea that your add-in is both an OAuth resource, relative to the Office host, and an OAuth client, relative to Microsoft Graph.</span></span>

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    
    ```

   > [!NOTE]
   > <span data-ttu-id="24359-166">Оставьте остальные параметры в разделе **appSettings** без изменений.</span><span class="sxs-lookup"><span data-stu-id="24359-166">Leave the other settings in the **appSettings** section unchanged.</span></span>

1. <span data-ttu-id="24359-167">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="24359-167">Save and close the file.</span></span>

1. <span data-ttu-id="24359-168">В проекте надстройки откройте файл манифеста Office-Add-in-ASPNET-SSO.xml.</span><span class="sxs-lookup"><span data-stu-id="24359-168">In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml”.</span></span>

1. <span data-ttu-id="24359-169">Перейдите в конец кода файла.</span><span class="sxs-lookup"><span data-stu-id="24359-169">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="24359-170">Над закрывающим тегом `</VersionOverrides>` вы найдете следующую часть кода:</span><span class="sxs-lookup"><span data-stu-id="24359-170">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

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

1. <span data-ttu-id="24359-171">Замените заполнитель {application_GUID here} *в обоих местах* кода идентификатором приложения, скопированным во время регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-171">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="24359-172">"{}" не являются частью идентификатора, поэтому не включайте их.</span><span class="sxs-lookup"><span data-stu-id="24359-172">The "{}" are not part of the ID, so do not include them.</span></span> <span data-ttu-id="24359-173">Это тот же идентификатор, который использовался для ClientID и Audience в файле web.config.</span><span class="sxs-lookup"><span data-stu-id="24359-173">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="24359-174">Значение **Resource** представляет собой **URI идентификатора приложения**, который вы задали, когда добавляли платформу веб-API при регистрации надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-174">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="24359-175">Раздел **Scopes** используется для создания диалогового окна предоставления разрешений, только если надстройка продается в AppSource.</span><span class="sxs-lookup"><span data-stu-id="24359-175">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="24359-176">Откройте вкладку **Предупреждения** в **списке ошибок** в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="24359-176">Open the **Warnings** tab of the **Error List** in Visual Studio.</span></span> <span data-ttu-id="24359-177">Предупреждение о том, что `<WebApplicationInfo>` не является действительным потомком `<VersionOverrides>`, означает, что используемая вами версия Visual Studio 2017 Preview не может распознать разметку единого входа.</span><span class="sxs-lookup"><span data-stu-id="24359-177">If there is a warning that `<WebApplicationInfo>` is not a valid child of `<VersionOverrides>`, your version of Visual Studio 2017 Preview does not  recognize the SSO markup.</span></span> <span data-ttu-id="24359-178">В качестве обходного решения в надстройке Word, Excel или PowerPoint можно выполнить указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="24359-178">As a workaround, do the following for a Word, Excel, or PowerPoint add-in.</span></span> <span data-ttu-id="24359-179">Если вы работаете с надстройкой Outlook, вы найдете решение ниже.</span><span class="sxs-lookup"><span data-stu-id="24359-179">(If you are working with an Outlook add-in see the workaround below.)</span></span>

   - <span data-ttu-id="24359-180">**Обходное решение для Word, Excel и Powerpoint**</span><span class="sxs-lookup"><span data-stu-id="24359-180">**Workaround for Word, Excel, and Powerpoint**</span></span>

        1. <span data-ttu-id="24359-181">Закомментируйте раздел `<WebApplicationInfo>` в манифесте прямо перед завершением узла `</VersionOverrides>`.</span><span class="sxs-lookup"><span data-stu-id="24359-181">Comment out the `<WebApplicationInfo>` section from the manifest just above the end of `</VersionOverrides>`.</span></span>

        2. <span data-ttu-id="24359-p116">Нажмите клавишу F5, чтобы запустить сеанс отладки. В результате будет создана копия манифеста в следующей папке (доступ к которой проще получить в **проводнике**, чем в Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span><span class="sxs-lookup"><span data-stu-id="24359-p116">Press F5 to start a debugging session. This will create a copy of the manifest in the following folder (which is easier to access in **File Explorer** than in Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span></span>

        3. <span data-ttu-id="24359-184">В копии манифеста удалите синтаксис комментария для раздела `<WebApplicationInfo>`.</span><span class="sxs-lookup"><span data-stu-id="24359-184">In the copy of the manifest, remove the comment syntax around the `<WebApplicationInfo>` section.</span></span>

        4. <span data-ttu-id="24359-185">Сохраните копию манифеста.</span><span class="sxs-lookup"><span data-stu-id="24359-185">Save the copy of the manifest.</span></span>

        5. <span data-ttu-id="24359-p117">Теперь необходимо принять меры, чтобы Visual Studio не перезаписал копию манифеста, когда вы в следующий раз нажмете клавишу F5. Щелкните правой кнопкой мыши узел решения в верхней части **обозревателя решений** (но не узлы проектов).</span><span class="sxs-lookup"><span data-stu-id="24359-p117">Now you must prevent Visual Studio from overwriting the copy of the manifest the next time you press F5. Right-click the solution node at the very top of **Solution Explorer** (not either of the project nodes).</span></span>

        6. <span data-ttu-id="24359-188">В контекстном меню выберите **Свойства**. Откроется диалоговое окно **Страницы свойств решения**.</span><span class="sxs-lookup"><span data-stu-id="24359-188">Select **Properties** from the context menu and a **Solution Property Pages** dialog box opens.</span></span>

        7. <span data-ttu-id="24359-189">Разверните пункт **Свойства конфигурации** и щелкните **Конфигурация**.</span><span class="sxs-lookup"><span data-stu-id="24359-189">Expand **Configuration Properties** and select **Configuration**.</span></span>

        8. <span data-ttu-id="24359-190">Снимите флажки **Выполнить сборку** и **Развернуть** в строке для проекта **Office-Add-in-ASPNET-SSO** (но *не* проекта **Office-Add-in-ASPNET-SSO-WebAPI**).</span><span class="sxs-lookup"><span data-stu-id="24359-190">Deselect **Build** and **Deploy** in the row for the **Office-Add-in-ASPNET-SSO** project (*not* the **Office-Add-in-ASPNET-SSO-WebAPI** project).</span></span>

        9. <span data-ttu-id="24359-191">Закройте диалоговое окно, нажав кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="24359-191">Press **OK** to close the dialog box.</span></span>

   - <span data-ttu-id="24359-192">**Обходное решение для Outlook**</span><span class="sxs-lookup"><span data-stu-id="24359-192">**Workaround for Outlook**</span></span>

        1. <span data-ttu-id="24359-193">Найдите файл `MailAppVersionOverridesV1_1.xsd` на компьютере, используемом для разработки.</span><span class="sxs-lookup"><span data-stu-id="24359-193">On your development machine, locate the existing `MailAppVersionOverridesV1_1.xsd`.</span></span> <span data-ttu-id="24359-194">Он должен находиться в том каталоге, в котором установлена среда Visual Studio, в папке `./Xml/Schemas/{lcid}`.</span><span class="sxs-lookup"><span data-stu-id="24359-194">This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`.</span></span> <span data-ttu-id="24359-195">Например, при обычной установке 32-разрядной версии VS 2017 в системе, где используется английский язык (США), полный путь будет выглядеть так: `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span><span class="sxs-lookup"><span data-stu-id="24359-195">For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span></span>

        2. <span data-ttu-id="24359-196">Измените имя существующего файла на `MailAppVersionOverridesV1_1.old`.</span><span class="sxs-lookup"><span data-stu-id="24359-196">Rename the existing file to `MailAppVersionOverridesV1_1.old`.</span></span>

        3. <span data-ttu-id="24359-197">Скопируйте измененную версию файла в папку: [Измененная схема MailAppVersionOverrides](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span><span class="sxs-lookup"><span data-stu-id="24359-197">Copy this modified version of the file into the folder: [Modified MailAppVersionOverrides Schema](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span></span>

1. <span data-ttu-id="24359-198">Сохраните и закройте главный файл манифеста в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="24359-198">Save and close the main manifest file in Visual Studio.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="24359-199">Код на стороне клиента</span><span class="sxs-lookup"><span data-stu-id="24359-199">Code the client side</span></span>

1. <span data-ttu-id="24359-p119">Откройте файл Home.js в папке **Scripts**. В нем уже есть следующий код:</span><span class="sxs-lookup"><span data-stu-id="24359-p119">Open the Home.js file in the **Scripts** folder. It already has some code in it:</span></span>
    * <span data-ttu-id="24359-202">Назначение методу `Office.initialize`, которое, в свою очередь, назначает обработчик события для нажатия кнопки `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="24359-202">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="24359-203">Метод `showResult` для отображения сообщения об ошибке (или данных, возвращаемых из Microsoft Graph) в нижней части области задач.</span><span class="sxs-lookup"><span data-stu-id="24359-203">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="24359-204">Метод `logErrors` для регистрации в консоли ошибок, которые не предназначены для пользователя.</span><span class="sxs-lookup"><span data-stu-id="24359-204">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

1. <span data-ttu-id="24359-p120">После назначения для метода `Office.initialize` добавьте приведенный ниже код. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p120">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="24359-207">При обработке ошибок в надстройке иногда автоматически выполняется еще одна попытка получить маркер доступа с помощью другого набора параметров.</span><span class="sxs-lookup"><span data-stu-id="24359-207">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="24359-208">Переменная счетчика `timesGetOneDriveFilesHasRun` и переменная флажка `triedWithoutForceConsent` используются, чтобы предотвратить циклическое повторение неудачных попыток получить маркер.</span><span class="sxs-lookup"><span data-stu-id="24359-208">The counter variable `timesGetOneDriveFilesHasRun`, and the flag variable `triedWithoutForceConsent` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span> 
    * <span data-ttu-id="24359-p122">Метод `getDataWithToken` создается на следующем шаге. Обратите внимание на то, что он присваивает параметру `forceConsent` значение `false`. Дополнительные сведения см. в описании следующего шага.</span><span class="sxs-lookup"><span data-stu-id="24359-p122">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. <span data-ttu-id="24359-p123">Под методом `getOneDriveFiles` добавьте приведенный ниже код. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p123">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="24359-p124">[getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) — это новый API в Office.js, позволяющий надстройке запрашивать у ведущего приложения Office (Excel, PowerPoint, Word и т. д.) маркер доступа к надстройке для пользователя, вошедшего в Office. Ведущее приложение Office, в свою очередь, запрашивает маркер у конечной точки Azure AD 2.0. Так как вы предварительно авторизовали ведущее приложение Office для надстройки во время ее регистрации, Azure AD отправит маркер.</span><span class="sxs-lookup"><span data-stu-id="24359-p124">The [](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office). The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token. Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="24359-216">Если вход в Office не выполнен, ведущее приложение Office предложит пользователю войти.</span><span class="sxs-lookup"><span data-stu-id="24359-216">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="24359-217">Параметр настроек задает для `forceConsent` значение `false`, поэтому пользователю не будет предлагаться разрешить ведущему приложению Office доступ к надстройке при каждом ее использовании.</span><span class="sxs-lookup"><span data-stu-id="24359-217">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in.</span></span> <span data-ttu-id="24359-218">При первом запуске надстройки вызов `getAccessTokenAsync` не будет выполнен, но логика обработки ошибок, которую вы добавите на следующем этапе, автоматически выполнит повторный вызов, при этом параметру `forceConsent` будет задано значение `true`, и пользователю будет предложено согласиться. Такая процедура выполняется только в первый раз.</span><span class="sxs-lookup"><span data-stu-id="24359-218">The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="24359-219">Вы создадите метод `handleClientSideErrors` позже.</span><span class="sxs-lookup"><span data-stu-id="24359-219">You will create the `handleClientSideErrors` method in a later step.</span></span>

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

1. <span data-ttu-id="24359-p126">Замените строку TODO1 на приведенные ниже строки. Метод `getData` и серверный маршрут /api/values создаются позже. Для конечной точки используется относительный URL-адрес, так как она должна размещаться на том же домене, что и надстройка.</span><span class="sxs-lookup"><span data-stu-id="24359-p126">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="24359-p127">Под методом `getOneDriveFiles` добавьте приведенный ниже код. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p127">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="24359-p128">Этот метод вызывает указанную конечную точку веб-API и передает ей тот же маркер доступа, который ведущее приложение Office использовало для доступа к надстройке. На стороне сервера этот маркер доступа будет использоваться в потоке "от имени" для получения маркера доступа к Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="24359-p128">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="24359-227">Вы создадите метод `handleServerSideErrors` позже.</span><span class="sxs-lookup"><span data-stu-id="24359-227">You will create the `handleServerSideErrors` method in a later step.</span></span>

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

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="24359-228">Создание методов обработки ошибок</span><span class="sxs-lookup"><span data-stu-id="24359-228">Create the error-handling methods</span></span>

1. <span data-ttu-id="24359-229">Под методом `getData` добавьте приведенный ниже метод.</span><span class="sxs-lookup"><span data-stu-id="24359-229">Below the `getData` method, add the following method.</span></span> <span data-ttu-id="24359-230">Этот метод будет обрабатывать ошибки в клиенте надстройки, когда ведущее приложение Office не сможет получить маркер доступа к веб-службе надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-230">This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service.</span></span> <span data-ttu-id="24359-231">Сообщения о таких ошибках содержат код ошибки, поэтому данный метод различает их с помощью оператора `switch`.</span><span class="sxs-lookup"><span data-stu-id="24359-231">These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Micrososoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="24359-232">Замените `TODO2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="24359-232">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="24359-233">Ошибка 13001 возникает, если пользователь не выполнил вход или без отклика отменил запрос на предоставление 2-го фактора проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="24359-233">Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor.</span></span> <span data-ttu-id="24359-234">В обоих случаях код повторно выполняет метод `getDataWithToken` и задает параметр для принудительного запрашивания входа.</span><span class="sxs-lookup"><span data-stu-id="24359-234">In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="24359-235">Замените `TODO3` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="24359-235">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="24359-236">Ошибка 13002 возникает, когда вход или предоставление разрешений прерывается.</span><span class="sxs-lookup"><span data-stu-id="24359-236">Error 13002 occurs when user's sign-in or consent was aborted.</span></span> <span data-ttu-id="24359-237">Попросите пользователя повторить попытку, но не более одного раза.</span><span class="sxs-lookup"><span data-stu-id="24359-237">Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. <span data-ttu-id="24359-p132">Замените `TODO4` приведенным ниже кодом. Ошибка 13003 возникает, когда пользователь вошел в систему с учетной записью, которая не является ни рабочей, ни учебной, ни учетной записью Майкрософт. Попросите пользователя выйти из системы, а затем снова войти, используя поддерживаемый тип учетной записи.</span><span class="sxs-lookup"><span data-stu-id="24359-p132">Replace `TODO4` with the following code. Error 13003 occurs when user is logged in with an account that is neither work or school, nor Micrososoft Account. Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > <span data-ttu-id="24359-241">Ошибки 13004 и 13005 не обрабатываются при использовании этого метода, так как они должны возникать только на стадии разработки.</span><span class="sxs-lookup"><span data-stu-id="24359-241">Errors 13004 and 13005 are not handled in this method because they should only occur in development.</span></span> <span data-ttu-id="24359-242">Их невозможно исправить с помощью кода среды выполнения, поэтому нет смысла сообщать о них пользователю.</span><span class="sxs-lookup"><span data-stu-id="24359-242">They cannot be fixed by runtime code and there would be no point in reporting them to an end user.</span></span>

1. <span data-ttu-id="24359-p134">Замените `TODO5` приведенным ниже кодом. Ошибка 13006 возникает, если происходит неопределенная ошибка ведущего приложения Office, которая может свидетельствовать о его нестабильном состоянии. Попросите пользователя перезапустить Office.</span><span class="sxs-lookup"><span data-stu-id="24359-p134">Replace `TODO5` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. <span data-ttu-id="24359-246">Замените `TODO6` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="24359-246">Replace `TODO6` with the following code.</span></span> <span data-ttu-id="24359-247">Ошибка 13007 возникает, когда нарушается взаимодействие ведущего приложения Office с AAD, из-за чего это приложение не может получить маркер доступа к веб-службе/приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-247">Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application.</span></span> <span data-ttu-id="24359-248">Это может быть из-за временного сбоя сети.</span><span class="sxs-lookup"><span data-stu-id="24359-248">This may be a temporary network issue.</span></span> <span data-ttu-id="24359-249">Попросите пользователя повторить попытку позже.</span><span class="sxs-lookup"><span data-stu-id="24359-249">Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. <span data-ttu-id="24359-p136">Замените `TODO7` приведенным ниже кодом. Ошибка 13008 возникает, когда пользователь запускает операцию, которая вызывает `getAccessTokenAsync`, до завершения предыдущего вызова.</span><span class="sxs-lookup"><span data-stu-id="24359-p136">Replace `TODO7` with the following code. Error 13008 occurs when the user tiggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. <span data-ttu-id="24359-252">Замените `TODO8` указанным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="24359-252">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="24359-253">Ошибка 13009 возникает, если надстройка не поддерживает принудительное запрашивание разрешения, но выполняется вызов `getAccessTokenAsync` с установкой для параметра `forceConsent` значения `true`.</span><span class="sxs-lookup"><span data-stu-id="24359-253">Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`.</span></span> <span data-ttu-id="24359-254">Обычно в таком случае код должен автоматически повторно запустить метод `getAccessTokenAsync` с параметром, имеющим значение `false`.</span><span class="sxs-lookup"><span data-stu-id="24359-254">In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`.</span></span> <span data-ttu-id="24359-255">Но в некоторых случаях вызов метода с установкой для параметра `forceConsent` значения `true` сам по себе является автоматическим откликом на ошибку вызова метода с установкой для параметра значения `false`.</span><span class="sxs-lookup"><span data-stu-id="24359-255">However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`.</span></span> <span data-ttu-id="24359-256">В этом случае код должен не повторять попытку, а предложить пользователю выйти и войти заново.</span><span class="sxs-lookup"><span data-stu-id="24359-256">In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. <span data-ttu-id="24359-257">Замените `TODO9` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="24359-257">Replace `TODO9` with the following code.</span></span>

    ```javascript
    default:
        logError(result);
        break;
    ```  


1. <span data-ttu-id="24359-p138">Под методом `handleClientSideErrors` добавьте приведенный ниже метод. Этот метод обрабатывает ошибки в веб-службе надстройки при неправильном выполнении потока "от имени" или получении данных от Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="24359-p138">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Parse the JSON response.

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO14: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. <span data-ttu-id="24359-260">Замените `TODO10` указанным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="24359-260">Replace `TODO10` with the following code.</span></span> <span data-ttu-id="24359-261">Обратите внимание, что для большинства ошибок `4xx`, которые веб-служба будет передавать клиентской части надстройки, в ответе будет свойство **ExceptionMessage**, содержащее номер ошибки AADSTS и другие данные.</span><span class="sxs-lookup"><span data-stu-id="24359-261">Note that for most of the `4xx` errors that the add-in's web service will pass to the add-in's client-side, there will be an **ExceptionMessage** property in the response that contains the AADSTS (Azure Active Directory Secure Token Service) error number as well as other data.</span></span> <span data-ttu-id="24359-262">Однако, когда AAD отправляет веб-службе надстройки запрос дополнительной проверки подлинности, этот запрос содержит специальное свойство **Claims** с кодом необходимой дополнительной проверки.</span><span class="sxs-lookup"><span data-stu-id="24359-262">However, when AAD sends a message to the add-in's web service asking for an additonal authentication factor, the message contains a special **Claims** property that specifies (with a code number) what additional factor is needed.</span></span> <span data-ttu-id="24359-263">API ASP.NET, которые создают и отправляют HTTP-ответы клиентам, не знают об этом свойстве **Claims**, поэтому не включают его в ответ.</span><span class="sxs-lookup"><span data-stu-id="24359-263">The ASP.NET APIs that create and send HTTP Responses to clients do not know about this **Claims** property, so they do not include it in the Response object.</span></span> <span data-ttu-id="24359-264">Серверный код, который вы создадите позже, будет вручную добавлять значение **Claims** в ответ, чтобы решить эту проблему.</span><span class="sxs-lookup"><span data-stu-id="24359-264">Server-side code that you will create in a later step will cope with this by manually adding the **Claims** value to the Response object.</span></span> <span data-ttu-id="24359-265">Это значение будет находиться в свойстве **Message**, поэтому код также должен анализировать это свойство.</span><span class="sxs-lookup"><span data-stu-id="24359-265">This value will be in the **Message** property, so the code needs to parse out that property as well.</span></span>

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. <span data-ttu-id="24359-p140">Замените `TODO11` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p140">Replace `TODO11` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="24359-268">Ошибка 50076 возникает, когда Microsoft Graph требует дополнительной проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="24359-268">Error 50076 occurs when Microsoft Graph requires an additional form of authentication.</span></span>
    * <span data-ttu-id="24359-269">Основное приложение Office должно получить новый маркер со значением **Claims** в качестве параметра `authChallenge`.</span><span class="sxs-lookup"><span data-stu-id="24359-269">The Office host should get a new token with the **Claims** value as the `authChallenge` option.</span></span> <span data-ttu-id="24359-270">В результате AAD предложит пользователю пройти все необходимые проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="24359-270">This tells AAD to prompt the user for all required forms of authentication.</span></span> 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }    
    ```

1. <span data-ttu-id="24359-p142">Замените `TODO12` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p142">Replace `TODO12` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="24359-273">Ошибка 65001 означает, что доступ к Microsoft Graph не был предоставлен (или был отозван) для одного или нескольких разрешений.</span><span class="sxs-lookup"><span data-stu-id="24359-273">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span> 
    * <span data-ttu-id="24359-274">Надстройка должна получить новый маркер с параметром `forceConsent`, имеющим значение `true`.</span><span class="sxs-lookup"><span data-stu-id="24359-274">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
       // getDataWithToken({ forceConsent: true });
    }    
    ```

1. <span data-ttu-id="24359-p143">Замените `TODO13` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p143">Replace `TODO13` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="24359-p144">Ошибка 70011 имеет несколько значений. Главное для этой надстройки — запрашивание недопустимого разрешения, поэтому код проверяет наличие полного описания ошибки, а не только номера.</span><span class="sxs-lookup"><span data-stu-id="24359-p144">Error 70011 has multiple meanings. The one that matters to this add-in is when it means that an invalid scope (permission) has been requested, so the code checks for the full error description, not just the number.</span></span>
    * <span data-ttu-id="24359-279">Надстройка должна сообщить об ошибке.</span><span class="sxs-lookup"><span data-stu-id="24359-279">The add-in should report the error.</span></span>

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }    
    ```

1. <span data-ttu-id="24359-p145">Замените `TODO14` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p145">Replace `TODO14` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="24359-282">Серверный код, который вы создадите позже, отправит сообщение `Missing access_as_user`, если разрешения `access_as_user` не будет в маркере доступа, который клиент надстройки отправит в AAD для использования в потоке "от имени".</span><span class="sxs-lookup"><span data-stu-id="24359-282">Server-side code that you create in a later step will send the message `Missing access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="24359-283">Надстройка должна сообщить об ошибке.</span><span class="sxs-lookup"><span data-stu-id="24359-283">The add-in should report the error.</span></span>

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }    
    ```

1. <span data-ttu-id="24359-p146">Замените `TODO15` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p146">Replace `TODO15` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="24359-p147">Библиотека идентификации, которую вы будете использовать в серверном коде (MSAL), должна предотвращать отправку в Microsoft Graph устаревших и недействительных маркеров. Но если это все-таки произойдет, Microsoft Graph вернет веб-службе надстройки ошибку с кодом `InvalidAuthenticationToken`. Серверный код, который вы создадите позже, передаст это сообщение клиенту надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-p147">The identity library that you will be using in the server-side code (Microsoft Authentication Library - MSAL) should ensure that no expired or invalid token is sent to Microsoft Graph; but if it does happen, the error that is returned to the add-in's web service from Microsoft Graph has the code `InvalidAuthenticationToken`. Server-side code you will create in a latter step will relay this message to the add-in's client.</span></span>
    * <span data-ttu-id="24359-288">В этом случае надстройка должна начать весь процесс проверки подлинности путем сброса переменных счетчика и флажка, а затем повторно вызвать метод обработчика кнопки.</span><span class="sxs-lookup"><span data-stu-id="24359-288">In this case, the add-in should start the entire authentication process over by resetting the counter and flag varibles, and then re-calling the button handler method.</span></span>

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }    
    ```

1. <span data-ttu-id="24359-289">Замените `TODO16` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="24359-289">Replace `TODO16` with the following code.</span></span>

    ```javascript
    else {
        logError(result);
    }    
    ```

1. <span data-ttu-id="24359-290">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="24359-290">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="24359-291">Код на стороне сервера</span><span class="sxs-lookup"><span data-stu-id="24359-291">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="24359-292">Настройка ПО промежуточного слоя OWIN</span><span class="sxs-lookup"><span data-stu-id="24359-292">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="24359-293">Откройте файл Startup.cs в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="24359-293">Open the Startup.cs file in the root of the project.</span></span>

1. <span data-ttu-id="24359-p148">Добавьте ключевое слово `partial` в объявление класса Startup, если его там еще нет. Оно должно выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="24359-p148">Add the keyword `partial` to the declaration of the Startup class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="24359-p149">Добавьте приведенную ниже строку в текст метода `Configuration`. Метод `ConfigureAuth` создается позже.</span><span class="sxs-lookup"><span data-stu-id="24359-p149">Add the following line to the body of the `Configuration` method. You create the `ConfigureAuth` method in a later step.</span></span>

    `ConfigureAuth(app);`

1. <span data-ttu-id="24359-298">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="24359-298">Save and close the file.</span></span>

1. <span data-ttu-id="24359-299">Щелкните правой кнопкой мыши папку **App_Start** и выберите **Добавить > Класс**.</span><span class="sxs-lookup"><span data-stu-id="24359-299">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="24359-300">В диалоговом окне **Добавить новый элемент** введите имя файла **Startup.Auth.cs** и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="24359-300">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="24359-301">Сократите имя пространства имен в новом файле до `Office_Add_in_ASPNET_SSO_WebAPI`.</span><span class="sxs-lookup"><span data-stu-id="24359-301">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="24359-302">Убедитесь, что в начале файла есть все приведенные ниже операторы `using`.</span><span class="sxs-lookup"><span data-stu-id="24359-302">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="24359-p150">Добавьте ключевое слово `partial` в объявление класса `Startup`, если его там еще нет. Оно должно выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="24359-p150">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="24359-p151">Добавьте приведенный ниже метод в класс `Startup`. Этот метод указывает, как ПО промежуточного слоя OWIN будет проверять маркеры доступа, передаваемые ему из метода `getData` в файле Home.js на стороне клиента. Процесс вызывается при каждом вызове конечной точки веб-API, содержащей атрибут `[Authorize]`.</span><span class="sxs-lookup"><span data-stu-id="24359-p151">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. <span data-ttu-id="24359-308">Замените TODO3 приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="24359-308">Replace the TODO3 with the following.</span></span> <span data-ttu-id="24359-309">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-309">Note about this code:</span></span>

    * <span data-ttu-id="24359-310">Код сообщает OWIN о необходимости убедиться, что аудитория и поставщик маркера, указанные в маркере доступа из ведущего приложения Office (который передается путем вызова метода `getData` на стороне клиента), должны совпадать со значениями, указанными в файле web.config.</span><span class="sxs-lookup"><span data-stu-id="24359-310">The code instructs OWIN to ensure that the audience and token issuer specified in the access token that comes from the Office host (and is passed on by the client-side call of `getData`) must match the values specified in the web.config.</span></span>
    * <span data-ttu-id="24359-p153">Если задать для свойства `SaveSigninToken` значение `true`, OWIN сохранит необработанный маркер из ведущего приложения Office. Он необходим надстройке, чтобы получить маркер доступа к Microsoft Graph в потоке "от имени".</span><span class="sxs-lookup"><span data-stu-id="24359-p153">Setting `SaveSigninToken` to `true` causes OWIN to save the raw token from the Office host. The add-in needs it to obtain an access token to Microsoft Graph with the “on behalf of” flow.</span></span>
    * <span data-ttu-id="24359-p154">ПО промежуточного слоя OWIN не проверяет разрешения. Разрешения маркера доступа, которые должны включать `access_as_user`, проверяются в контроллере.</span><span class="sxs-lookup"><span data-stu-id="24359-p154">Scopes are not validated by the OWIN middleware. The scopes of the access token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. <span data-ttu-id="24359-p155">Замените TODO4 приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p155">Replace TODO4 with the following. Note about this code:</span></span>

    * <span data-ttu-id="24359-317">Метод `UseOAuthBearerAuthentication` вызывается вместо более распространенного метода `UseWindowsAzureActiveDirectoryBearerAuthentication`, так как последний несовместим с конечной точкой Azure AD версии 2.</span><span class="sxs-lookup"><span data-stu-id="24359-317">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="24359-318">ПО промежуточного слоя OWIN использует URL-адрес обнаружения, передаваемый методу, чтобы получить ключ, необходимый для проверки подписи в маркере доступа, полученном из ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="24359-318">The discovery URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the access token received from the Office host.</span></span>

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. <span data-ttu-id="24359-319">Сохраните и закройте файл.</span><span class="sxs-lookup"><span data-stu-id="24359-319">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="24359-320">Создание контроллера /api/values</span><span class="sxs-lookup"><span data-stu-id="24359-320">Create the /api/values controller</span></span>

1. <span data-ttu-id="24359-321">Откройте файл **Controllers\ValueController.cs**.</span><span class="sxs-lookup"><span data-stu-id="24359-321">Open the file **Controllers\ValueController.cs**.</span></span>

2. <span data-ttu-id="24359-322">Убедитесь, что в начале файла есть приведенные ниже инструкции с `using`.</span><span class="sxs-lookup"><span data-stu-id="24359-322">Ensure that the following `using` statements are at the top of the file.</span></span>

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

3. <span data-ttu-id="24359-p156">Над строкой с объявлением `ValuesController` добавьте атрибут `[Authorize]`. Это гарантирует, что надстройка будет выполнять процесс авторизации, настроенный в последней процедуре, при каждом вызове метода контроллера. Вызывать методы контроллера можно только при наличии действительного маркера доступа к надстройке.</span><span class="sxs-lookup"><span data-stu-id="24359-p156">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

    > [!NOTE]
    > <span data-ttu-id="24359-326">Производственная служба веб-API на основе ASP.NET MVC должна иметь специальную логику для потока "от имени" в одном или нескольких пользовательских классах [FilterAttribute](https://docs.microsoft.com/previous-versions/aspnet/web-frameworks/hh834645(v=vs.108)).</span><span class="sxs-lookup"><span data-stu-id="24359-326">A production ASP.NET MVC Web API service should have custom logic for the on-behalf-of flow in one or more custom [FilterAttribute](https://docs.microsoft.com/previous-versions/aspnet/web-frameworks/hh834645(v=vs.108)) classes.</span></span> <span data-ttu-id="24359-327">В этом примере логика помещается в главный контроллер, чтобы можно было легко проследить весь поток авторизации и логику получения данных.</span><span class="sxs-lookup"><span data-stu-id="24359-327">This educational sample puts the logic in the main controller so that the entire flow of the authorization and data fetching logic can be easily followed.</span></span> <span data-ttu-id="24359-328">Такая же модель используется в примерах авторизации в разделе [Azure Samples](https://github.com/Azure-Samples/).</span><span class="sxs-lookup"><span data-stu-id="24359-328">This also makes the sample consistent with the pattern of authorization samples in [Azure Samples](https://github.com/Azure-Samples/).</span></span>    

4. <span data-ttu-id="24359-329">Добавьте приведенный ниже метод в `ValuesController`.</span><span class="sxs-lookup"><span data-stu-id="24359-329">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="24359-330">Обратите внимание, что возвращаемое значение — `Task<HttpResponseMessage>`, а не `Task<IEnumerable<string>>`, которое чаще используется для метода `GET api/values`.</span><span class="sxs-lookup"><span data-stu-id="24359-330">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="24359-331">Это побочный эффект нахождения пользовательской логики авторизации в контроллере: при возникновении некоторых ошибок веб-служба должна отправлять HTTP-ответ клиенту надстройки.</span><span class="sxs-lookup"><span data-stu-id="24359-331">This is a side effect of that fact that our custom authorization logic will be in the controller: some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span> 

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

5. <span data-ttu-id="24359-332">Замените `TODO1` приведенным ниже кодом, чтобы убедиться, что в токене указано разрешение `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="24359-332">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span>

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
    > <span data-ttu-id="24359-p159">Для авторизации API, который отвечает за поток выполнения от имени другого субъекта, в случае надстроек Office используйте только область `access_as_user`. Для других API в службе должны быть предусмотрены отдельные требования, касающиеся областей. Это ограничивает объекты, к которым можно получить доступ с помощью маркеров, полученных набором приложений Office.</span><span class="sxs-lookup"><span data-stu-id="24359-p159">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.</span></span>

6. <span data-ttu-id="24359-p160">Замените `TODO2` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p160">Replace `TODO2` with the following code. Note about this code:</span></span>
    * <span data-ttu-id="24359-337">Код преобразует необработанный маркер доступа, полученный от ведущего приложения Office, в объект `UserAssertion`, который будет передан другому методу.</span><span class="sxs-lookup"><span data-stu-id="24359-337">It turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.</span></span>
    * <span data-ttu-id="24359-p161">Надстройка больше не выступает в роли ресурса (или аудитории), доступ к которому необходим ведущему приложению Office и пользователю. Теперь она сама является клиентом, которому необходим доступ к Microsoft Graph. `ConfidentialClientApplication` — это объект "контекста клиента" MSAL.</span><span class="sxs-lookup"><span data-stu-id="24359-p161">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="24359-p162">Третий параметр конструктора `ConfidentialClientApplication` — URL-адрес перенаправления. На самом деле он не используется в потоке "от имени", но все равно рекомендуется указывать правильный URL-адрес. С помощью четвертого и пятого параметров можно определить постоянное хранилище, которое позволяет повторно использовать действительные маркеры в разных сеансах с надстройкой. В этом примере не реализуется постоянное хранилище.</span><span class="sxs-lookup"><span data-stu-id="24359-p162">The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the “on behalf of” flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.</span></span>
    * <span data-ttu-id="24359-344">Для работы библиотеки MSAL требуются области `openid` и `offline_access`, но если код их избыточно запрашивает, возникает ошибка.</span><span class="sxs-lookup"><span data-stu-id="24359-344">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them.</span></span> <span data-ttu-id="24359-345">Кроме того, ошибка возникнет, если код запросит `profile` (фактически используется только при получении ведущим приложением Office токена для веб-приложения надстройки).</span><span class="sxs-lookup"><span data-stu-id="24359-345">It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application.</span></span> <span data-ttu-id="24359-346">Поэтому явным образом запрашивается только `Files.Read.All`.</span><span class="sxs-lookup"><span data-stu-id="24359-346">So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. <span data-ttu-id="24359-p164">Замените `TODO3` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p164">Replace `TODO3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="24359-p165">Для начала метод `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` проверит кэш MSAL, который находится в памяти, на наличие подходящего маркера доступа. Только в случае его отсутствия запускается поток "от имени" с конечной точкой Azure AD версии 2.</span><span class="sxs-lookup"><span data-stu-id="24359-p165">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="24359-351">Если ресурс Microsoft Graph требует многофакторной проверки подлинности, а пользователь еще не предоставил соответствующие данные, AAD вызовет исключение, содержащее свойство Claims.</span><span class="sxs-lookup"><span data-stu-id="24359-351">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will throw an exception containing a Claims property.</span></span>
    * <span data-ttu-id="24359-p166">Значение свойства Claims необходимо передать клиенту, который передаст его ведущему приложению Office. Последнее добавит его в запрос на получение нового токена. AAD предложит пользователю пройти все необходимые проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="24359-p166">The Claims property value must be passed to the client which will pass it to the Office host, which will then include it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="24359-354">Любые исключения, отличные от типа `MsalServiceException`, не перехватываются преднамеренно, поэтому будут переданы клиенту в виде сообщений `500 Server Error`.</span><span class="sxs-lookup"><span data-stu-id="24359-354">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

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

8. <span data-ttu-id="24359-p167">Замените `TODO3a` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p167">Replace `TODO3a` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="24359-p168">Если ресурс Microsoft Graph требует многофакторной проверки подлинности, а пользователь еще не предоставил соответствующие данные, AAD вернет состояние "400 Bad Request" с ошибкой AADSTS50076 и свойство **Claims**. MSAL выдает исключение **MsalUiRequiredException** (наследуется от исключения **MsalServiceException**) с этой информацией.</span><span class="sxs-lookup"><span data-stu-id="24359-p168">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will return "400 Bad Request" with error AADSTS50076 and a **Claims** property. MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span> 
    * <span data-ttu-id="24359-p169">Значение свойства **Claims** необходимо передать клиенту, который передаст его ведущему приложению Office. Последнее добавит его в запрос на получение нового токена. AAD предложит пользователю пройти все необходимые проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="24359-p169">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="24359-361">API, которые создают HTTP-ответы из исключений, не знают о свойстве **Claims**, поэтому не включают его в ответ.</span><span class="sxs-lookup"><span data-stu-id="24359-361">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object.</span></span> <span data-ttu-id="24359-362">Нам нужно создать сообщение с ним вручную.</span><span class="sxs-lookup"><span data-stu-id="24359-362">We have to manually create a message that includes it.</span></span> <span data-ttu-id="24359-363">Однако настраиваемое свойство **Message** блокирует создание свойства **ExceptionMessage**, поэтому единственный способ передать идентификатор ошибки `AADSTS50076` клиенту — добавить его в настраиваемое свойство **Message**.</span><span class="sxs-lookup"><span data-stu-id="24359-363">A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**.</span></span> <span data-ttu-id="24359-364">Код JavaScript в клиенте должен будет определить, какое свойство содержится в ответе (**Message** или **ExceptionMessage**).</span><span class="sxs-lookup"><span data-stu-id="24359-364">JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="24359-365">Сообщение создается в формате JSON, чтобы клиентский код JavaScript мог проанализировать его с помощью известных методов объекта `JSON`.</span><span class="sxs-lookup"><span data-stu-id="24359-365">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known `JSON` object methods.</span></span>
    * <span data-ttu-id="24359-366">Вы создадите метод `SendErrorToClient` позже.</span><span class="sxs-lookup"><span data-stu-id="24359-366">You will create the `SendErrorToClient` method in a later step.</span></span> <span data-ttu-id="24359-367">Его второй параметр — объект **Exception**.</span><span class="sxs-lookup"><span data-stu-id="24359-367">It's second parameter is an **Exception** object.</span></span> <span data-ttu-id="24359-368">В этом случае код передает `null`, потому что включение объекта **Exception** блокирует включение свойства **Message** в создаваемый HTTP-ответ.</span><span class="sxs-lookup"><span data-stu-id="24359-368">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

9. <span data-ttu-id="24359-p172">Замените `TODO3b` и `TODO3c` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p172">Replace `TODO3b` and `TODO3c` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="24359-371">Если вызов AAD содержал по крайней мере одно разрешение, которое не предоставил ни пользователь, ни администратор клиента (или оно было отозвано),</span><span class="sxs-lookup"><span data-stu-id="24359-371">If the call to AAD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked).</span></span> <span data-ttu-id="24359-372">AAD вернет состояние "400 Bad Request" с ошибкой `AADSTS65001`.</span><span class="sxs-lookup"><span data-stu-id="24359-372">AAD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="24359-373">MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию.</span><span class="sxs-lookup"><span data-stu-id="24359-373">MSAL throws a **MsalUiRequiredException** with this information.</span></span> <span data-ttu-id="24359-374">Клиент должен вызвать метод `getAccessTokenAsync` повторно, используя параметр `{ forceConsent: true }`.</span><span class="sxs-lookup"><span data-stu-id="24359-374">The client should re-call `getAccessTokenAsync` with the option `{ forceConsent: true }`.</span></span>
    *  <span data-ttu-id="24359-375">Если вызов AAD содержал по крайней мере одно нераспознанное разрешение, AAD вернет состояние "400 Bad Request" с ошибкой `AADSTS70011`.</span><span class="sxs-lookup"><span data-stu-id="24359-375">If the call to AAD contained at least one scope that AAD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="24359-376">MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию.</span><span class="sxs-lookup"><span data-stu-id="24359-376">MSAL throws a **MsalUiRequiredException** with this information.</span></span> <span data-ttu-id="24359-377">Клиент должен сообщить об этом пользователю.</span><span class="sxs-lookup"><span data-stu-id="24359-377">The client should inform the user.</span></span>
    *  <span data-ttu-id="24359-378">Полное описание включается, так как ошибка 70011 возвращается и в других случаях, и ее следует обрабатывать в этой надстройке, только когда она означает запрос недопустимого разрешения.</span><span class="sxs-lookup"><span data-stu-id="24359-378">The entire description is included beause 70011 is returned in other conditions and we it should only be handled in this add-in when it means that there is an invalid scope.</span></span> 
    *  <span data-ttu-id="24359-p175">Объект **MsalUiRequiredException** передается методу `SendErrorToClient`. Это гарантирует, что свойство **ExceptionMessage**, содержащее информацию об ошибке, будет включено в HTTP-отклик.</span><span class="sxs-lookup"><span data-stu-id="24359-p175">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>
    *  <span data-ttu-id="24359-381">Сообщения нет, поэтому в качестве третьего параметра передается `null`.</span><span class="sxs-lookup"><span data-stu-id="24359-381">There is no custom message, so `null` is passed for the third parameter.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

10. <span data-ttu-id="24359-382">Замените `TODO3d` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="24359-382">Replace `TODO3d` with the following code.</span></span> <span data-ttu-id="24359-383">Обратите внимание, что код повторно выдает исключение, а не передает его в собственном HTTP-ответе с состоянием **HttpStatusCode.Forbidden** (401).</span><span class="sxs-lookup"><span data-stu-id="24359-383">Note that the code rethrows the exception instead of relaying it in a custom HTTP Response with **HttpStatusCode.Forbidden** (401).</span></span> <span data-ttu-id="24359-384">В результате ASP.NET отправляет собственный HTTP-ответ с состоянием "500 Server Error".</span><span class="sxs-lookup"><span data-stu-id="24359-384">The effect of this is that the ASP.NET will send its own HTTP Response with status "500 Server Error".</span></span>

    ```csharp
    else
    {
        throw e;
    }  
    ```

11. <span data-ttu-id="24359-p177">Замените `TODO4` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p177">Replace `TODO4` with the following. Note about this code:</span></span>

    * <span data-ttu-id="24359-p178">Классы `GraphApiHelper` и `ODataHelper` определяются в файлах из папки **Helpers**. Класс `OneDriveItem` определяется в файле из папки **Models**. В этой статье не представлено подробное описание этих классов, так как оно не имеет отношения к авторизации и единому входу.</span><span class="sxs-lookup"><span data-stu-id="24359-p178">The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.</span></span>
    * <span data-ttu-id="24359-390">Производительность будет выше, если запрашивать у Microsoft Graph только действительно необходимые данные, поэтому в коде заданы параметры ` $select` и `$top`. Первый из них показывает, что нужно только свойство name, второй — что требуются только первые три названия папок или файлов.</span><span class="sxs-lookup"><span data-stu-id="24359-390">Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a ` $select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first three folder or file names.</span></span>
    * <span data-ttu-id="24359-391">Если отправленный в Microsoft Graph токен недействителен, Microsoft Graph возвращает ошибку "401 Unauthorized" с кодом "InvalidAuthenticationToken".</span><span class="sxs-lookup"><span data-stu-id="24359-391">If the token sent to Microsoft Graph is invalid, Microsoft Graph sends a "401 Unauthorized" error with the code "InvalidAuthenticationToken".</span></span> <span data-ttu-id="24359-392">ASP.NET затем выдает исключение **RuntimeBinderException**.</span><span class="sxs-lookup"><span data-stu-id="24359-392">ASP.NET then throws a **RuntimeBinderException**.</span></span> <span data-ttu-id="24359-393">Это также происходит, когда срок действия токена истек, хотя MSAL должна предотвращать отправку таких токенов.</span><span class="sxs-lookup"><span data-stu-id="24359-393">This is also what happens when the token is expired, although MSAL should prevent that from ever happening.</span></span> 

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

12. <span data-ttu-id="24359-p180">Замените `TODO5` приведенным ниже кодом. Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-p180">Replace `TODO5` with the following. Note about this code:</span></span> 

    * <span data-ttu-id="24359-p181">Хотя приведенный выше код запрашивает только свойство *name* элементов OneDrive, Microsoft Graph всегда включает свойство *eTag* для элементов OneDrive. Чтобы сократить количество полезных данных, отправляемых клиенту, приведенный ниже код преобразует результаты, оставляя только имена элементов.</span><span class="sxs-lookup"><span data-stu-id="24359-p181">Although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.</span></span>
    * <span data-ttu-id="24359-398">Список из трех файлов и папок OneDrive отправляется клиенту в виде HTTP-ответа "200 OK".</span><span class="sxs-lookup"><span data-stu-id="24359-398">The list of three OneDrive files and folders is sent to the client as a "200 OK" HTTP Response.</span></span>

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

13. <span data-ttu-id="24359-399">Добавьте приведенный ниже метод под методом Get.</span><span class="sxs-lookup"><span data-stu-id="24359-399">Below the Get method, add the following method.</span></span> <span data-ttu-id="24359-400">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="24359-400">About this code note:</span></span>  

    * <span data-ttu-id="24359-401">Метод передает клиенту информацию об исключении на стороне сервера.</span><span class="sxs-lookup"><span data-stu-id="24359-401">The method relays to the client information about a server-side exception.</span></span> 
    * <span data-ttu-id="24359-402">Если методу будет передано исходное исключение, конструктор HttpError включит информацию из исключения в свойство **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="24359-402">If the original exception is passed to the method, then the HttpError constuctor will include information from the exception object in an **ExceptionMessage** property.</span></span>  
    * <span data-ttu-id="24359-403">Если в виде исключения будет передано значение `null`, конструктор HttpError включит параметр message в свойство **Message**. Свойства **ExceptionMessage** не будет.</span><span class="sxs-lookup"><span data-stu-id="24359-403">If `null` is passed for the exception, then the HttpError constuctor will include the message parameter in a **Message** property and there is no **ExceptionMessage** property.</span></span>

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

## <a name="run-the-add-in"></a><span data-ttu-id="24359-404">Запуск надстройки</span><span class="sxs-lookup"><span data-stu-id="24359-404">Run the add-in</span></span>

1. <span data-ttu-id="24359-405">Убедитесь в наличии нескольких файлов в OneDrive, чтобы можно было проверить результаты.</span><span class="sxs-lookup"><span data-stu-id="24359-405">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="24359-p183">В Visual Studio нажмите клавишу F5. Откроется PowerPoint, где на ленте **Главная** появится группа **SSO ASP.NET**.</span><span class="sxs-lookup"><span data-stu-id="24359-p183">In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.</span></span>

1. <span data-ttu-id="24359-408">Нажмите кнопку **Show Add-in** (Показать надстройку) в этой группе, чтобы увидеть пользовательский интерфейс надстройки в области задач.</span><span class="sxs-lookup"><span data-stu-id="24359-408">Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane.</span></span>

1. <span data-ttu-id="24359-p184">Нажмите кнопку **Get My Files from OneDrive** (Получить мои файлы из OneDrive). Если вы не вошли в Office, вам будет предложено войти.</span><span class="sxs-lookup"><span data-stu-id="24359-p184">Press the button **Get My Files from OneDrive**. If you are not signed into Office, you'll be prompted to sign in.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="24359-411">Если ранее вы вошли в Office, используя другой идентификатор, и не закрыли некоторые из открытых тогда приложений Office, Office может не сменить идентификатор (даже если кажется, что это сделано для PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="24359-411">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="24359-412">Если это произойдет, возможен сбой при вызове Microsoft Graph или возврат данных для другого идентификатора.</span><span class="sxs-lookup"><span data-stu-id="24359-412">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="24359-413">Чтобы избежать этого, *закройте все приложения Office*, прежде чем нажимать кнопку **Get My Files from OneDrive** (Получить мои файлы из OneDrive).</span><span class="sxs-lookup"><span data-stu-id="24359-413">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>

1. <span data-ttu-id="24359-p186">После входа под кнопкой появится список файлов и папок из OneDrive. Это может занять более 15 секунд, особенно в первый раз.</span><span class="sxs-lookup"><span data-stu-id="24359-p186">After you are signed in, a list of your files and folders on OneDrive will appear below the button. This may take over 15 seconds, especially the first time.</span></span>
