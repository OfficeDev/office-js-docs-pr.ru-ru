---
title: Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET
description: Пошаговое руководство по созданию (или преобразованию) надстройки Office с ASP.NET серверной частью для использования единого входа.
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: b0179429f9d81b893394278580b6ef8891dd0a87
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/28/2022
ms.locfileid: "68842105"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET

После того как пользователи войдут в Office, ваша надстройка сможет использовать те же учетные данные для предоставления им доступа к нескольким приложениям без необходимости повторного входа. Общие сведения см. в статье [Включение единого входа в надстройке Office](sso-in-office-add-ins.md).
В этой статье описывается процесс включения единого входа (SSO) в надстройке, созданной с помощью ASP.NET.

## <a name="prerequisites"></a>Предварительные требования

- Visual Studio 2019 или более поздней версии.

- Рабочая нагрузка **разработки Office/SharePoint** при настройке Visual Studio.

- [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

- По крайней мере несколько файлов и папок, хранящихся в OneDrive для бизнеса в подписке Microsoft 365.

- Учетная запись Azure с активной подпиской — [создайте учетную запись бесплатно](https://azure.microsoft.com/free/?WT.mc_id=A261C142F).

## <a name="set-up-the-starter-project"></a>Настройка начального проекта

Клонируйте или скачайте репозиторий [Office Add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).

> [!NOTE]
> Существует две версии примера.
>
> - The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.
> - Версия примера в папке **Complete** идентична надстройке, которую вы бы создали, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью. Чтобы использовать готовую версию, просто выполните действия, описанные в этой статье, но замените папку "Before" на папку "Complete" и пропустите разделы **Код на стороне клиента** и **Код на стороне сервера**.

Используйте следующие значения заполнителей для последующих шагов регистрации приложения.

| Заполнитель           | Значение                                           |
|-----------------------|-------------------------------------------------|
| `<add-in-name>`       | **Office-Add-in-ASPNET-SSO**                    |
| `<redirect-platform>` | **Web**                                         |
| `<redirect-uri>`      | `https://localhost:44355/AzureADAuth/Authorize` |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="configure-the-solution"></a>Настройка решения

1. В корне папки **Before** откройте SLN-файл решения в **Visual Studio**. В **обозревателе решений** щелкните правой кнопкой мыши верхний узел (узел решения, а не узлы проектов) и выберите **Назначить запускаемые проекты**.

1. В разделе **Общие свойства** выберите **Запускаемый проект**, а затем **Несколько запускаемых проектов**. Убедитесь, что для параметра **Действие** в обоих проектах установлено значение **Запуск** и что проект, заканчивающийся на "...WebAPI", указан в списке первым. Закройте диалоговое окно.

1. Вернитесь **в Обозреватель решений** выберите (не щелкайте правой кнопкой мыши) проект **Office-Add-in-ASPNET-SSO-WebAPI**. Откроется область **Свойства**. Убедитесь, что для параметра **SSL включен** задано значение **True**. Убедитесь, что **URL-адрес SSL** указан как `http://localhost:44355/`.

1. В файле web.config используйте значения, скопированные ранее. Для **ida:ClientID** и **ida:Audience** укажите **идентификатор приложения (клиента)**, для **ida:Password** — секрет клиента. Кроме того, задайте **для ida:Domain** значение `http://localhost:44355` (в конце нет косой черты "/").

    > [!NOTE]
    > **Идентификатор приложения (клиента)** — это значение аудитории, когда другие приложения, такие как клиентское приложение Office (например, PowerPoint, Word, Excel), ищут авторизованный доступ к приложению. Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.

1. Если вы не указали вариант "Учетные записи только в этом каталоге организации" для параметра **ПОДДЕРЖИВАЕМЫЕ ТИПЫ УЧЕТНЫХ ЗАПИСЕЙ** при регистрации настройки, сохраните и закройте файл web.config. В противном случае сохраните его, но оставьте открытым. 

1. В **Обозреватель решений** выберите проект **Office-Add-in-ASPNET-SSO** и откройте файл манифеста надстройки "Office-Add-in-ASPNET-SSO.xml", а затем прокрутите его вниз. Сразу над конечным `</VersionOverrides>` тегом вы найдете следующую разметку.

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

1. Замените заполнитель "$application_GUID here$" *в обоих местах* разметки идентификатором приложения, скопированным при регистрации надстройки. Символы "$" не входят в состав идентификатора, их не нужно вставлять. Это тот же идентификатор, который использовался для ClientID и Audience в файле web.config.

    > [!NOTE]
    > Значением **\<Resource\>** является **URI идентификатора приложения** , заданный при регистрации надстройки. Раздел **\<Scopes\>** используется только для создания диалогового окна согласия, если надстройка продается через AppSource.

1. Сохраните и закройте файл.

### <a name="setup-for-single-tenant"></a>Настройка в однотенантном режиме

Если вы выбрали "Учетные записи только в этом каталоге организации" для **параметра ПОДДЕРЖИВАЕМЫЕ ТИПЫ УЧЕТНЫх записей** при регистрации надстройки, необходимо выполнить следующие дополнительные действия по настройке.

1. Вернитесь на портал Azure и откройте колонку **Обзор** регистрации надстройки. Скопируйте **Идентификатор каталога (клиента)**.

1. В файле web.config замените "common" в значении **ida:Authority** на GUID, скопированный на предыдущем шаге.   После этого значение должно выглядеть следующим образом: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.

1. Сохраните и закройте файл web.config.

## <a name="code-the-client-side"></a>Код на стороне клиента

1. Откройте файл HomeES6.js в папке **Scripts**. В нем уже есть код.

    - Полизаполнение, которое назначает объект Office.Promise глобальному объекту window, чтобы надстройка могла работать, если в Office используется пользовательский интерфейс Internet Explorer. (Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).)
    - Назначение `Office.initialize` функции, которая, в свою очередь, назначает обработчик событию нажатия кнопки `getGraphAccessTokenButton` .
    - Метод `showResult` для отображения сообщения об ошибке (или данных, возвращаемых из Microsoft Graph) в нижней части области задач.
    - Метод `logErrors` для регистрации в консоли ошибок, которые не предназначены для пользователя.
    - Код, реализующий резервную систему авторизации, которую надстройка будет использовать в сценариях, где единый вход не поддерживается или произошла ошибка.

1. После назначения добавьте `Office.initialize`следующий код. Вот что нужно знать об этом коде:

    - При обработке ошибок в надстройке иногда автоматически выполняется еще одна попытка получить маркер доступа с помощью другого набора параметров. Переменная счетчика `retryGetAccessToken` используется, чтобы предотвратить циклическое повторение неудачных попыток получить маркер.
    - Функция `getGraphData` определяется ключевым словом `async` в ES6. Синтаксис ES6 значительно упрощает использование API единого входа в надстройках Office. Это единственный файл в решении, в котором используется синтаксис, не поддерживаемый в Internet Explorer. "ES6" включается в имя файла в качестве напоминания. Компилятор TSC используется в решении для компиляции этого файла в ES5, чтобы надстройка могла работать, если в Office используется пользовательский интерфейс Internet Explorer. (См. файл tsconfig.json в корневой папке проекта.)

    ```javascript
    let retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. `getGraphData` После функции добавьте следующую функцию. Обратите внимание, что функция `handleClientSideErrors` будет создана позже.

    > [!NOTE]
    > Чтобы отличить два маркера доступа, с которыми вы работаете в этой статье, маркер, возвращенный getAccessToken(), называется маркером начальной загрузки. Позже он обменивается через поток On-Behalf-Of для нового маркера с доступом к Microsoft Graph.

    ```javascript
    async function getDataWithToken(options) {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for a new access token to Microsoft Graph and then get the data
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


1. Замените `TODO 1` приведенным ниже кодом, чтобы получить маркер доступа от узла Office. Параметр *options* содержит следующие параметры, переданные из предыдущей `getGraphData()` функции.

    - `allowSignInPrompt` имеет значение true. При этом Office предложит пользователю войти, если он еще не вошел в Office.
    - `allowConsentPrompt` имеет значение true. При этом Office предложит пользователю дать согласие на предоставление надстройке доступа к Microsoft Azure Active Directory профилю пользователя, если согласие еще не предоставлено. (Результирующий запрос *не* позволяет пользователю предоставить согласие на какие-либо области Microsoft Graph.)
    - `forMSGraphAccess` имеет значение true. При этом Office возвращает ошибку (код 13012), если пользователь или администратор не предоставил согласие на использование областей Graph для надстройки. Чтобы получить доступ к Microsoft Graph, надстройка должна обменять маркер доступа на новый маркер доступа через поток on-behalf-of. Установка `forMSGraphAccess` значения true позволяет избежать сценария, в котором **getAccessToken()** выполняется успешно, но затем поток on-behalf-of позже завершается сбоем для Microsoft Graph. Код на стороне клиента может реагировать на ошибку 13012, переходя на резервную систему авторизации.

    Кроме того, обратите внимание на следующий код:

    - Вы создадите функцию `getData` позже.
    - Параметр `/api/values` представляет собой URL-адрес контроллера на стороне сервера, который будет использовать поток on-behalf-of для обмена маркером на новый маркер доступа для вызова Microsoft Graph.

    ```javascript
    let bootstrapToken = await Office.auth.getAccessToken(options);

    getData("/api/values", bootstrapToken);
    ```

1. `getGraphData` После функции добавьте следующее. Вот что нужно знать об этом коде:

    - Он используется и в системах единого входа, и в резервных системах авторизации.
    - Параметр `relativeUrl` является контроллером на стороне сервера.
    - Параметр `accessToken` может быть маркером начальной загрузки или маркером полного доступа.
    - `writeFileNamesToOfficeDocument` уже включен в проект.
    - Вы создадите функцию `handleServerSideErrors` позже.

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

### <a name="handle-client-side-errors"></a>Обработка ошибок на стороне клиента

1. `getData` После функции добавьте следующую функцию. Обратите внимание, что `error.code` — это число (обычно в диапазоне 13xxx).

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

1. Замените `TODO 2` приведенным ниже кодом. Дополнительные сведения об этих ошибках см. в статье [Устранение ошибок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md).

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-in's functions anyway. If you want to sign in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-in's functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
        break;
    case 13006:
        // Only seen in Office on the web.
        showResult(["Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. Замените `TODO 3` приведенным ниже кодом. Во всех других случаях надстройка переходит на резервную систему авторизации. Дополнительные сведения об этих ошибках см. [в статье Устранение неполадок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md). В этой надстройке резервная система открывает диалоговое окно, в котором требуется, чтобы пользователь выполнил вход, даже если пользователь уже выполнил вход.

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a>Обработка ошибок на стороне сервера

1. `handleClientSideErrors` После функции добавьте следующую функцию.

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. Замените `TODO 4` приведенным ниже кодом. Вот что нужно знать об этом коде: классы ошибок в ASP.NET были созданы до появления MFA. Побочным эффектом того, как логика на стороне сервера обрабатывает запросы второго фактора проверки подлинности, является то, что у ошибки на стороне сервера, отправляемой клиенту, есть свойство **Message**, но нет свойства **ExceptionMessage**. Однако у всех остальных ошибок будет свойство **ExceptionMessage**, поэтому клиентский код должен проанализировать ответ для обоих свойств.  Одна из переменных не будет определена.

    ```javascript
    const message = JSON.parse(result.responseText).Message;
    const exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. Замените `TODO 5` приведенным ниже кодом. Когда Microsoft Graph требует дополнительной проверки подлинности, он отправляет ошибку AADSTS50076. Она содержит сведения о дополнительном требовании в свойстве **Message.Claims**. Чтобы обработать эту ошибку, код делает вторую попытку получить маркер начальной загрузки, но в этот раз он включает запрос дополнительного фактора в виде значения параметра `authChallenge`, который предписывает Azure AD предложить пользователю пройти все требуемые проверки подлинности. 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            const claims = JSON.parse(message).Claims;
            const claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. Замените `TODO 6` следующим кодом:

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. Замените `TODO 7` приведенным ниже кодом. Обратите внимание, что иногда срок действия маркера начальной загрузки не истекает в момент его проверки в Office, но истекает ко времени его попадания в Azure AD для замены. Служба Azure AD ответит ошибкой AADSTS500133. В этом случае код вызывает API единого входа (но не более одного раза). На этот раз Office возвращает новый маркер начальной загрузки, срок действия которого не истек.  

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. Замените `TODO 8` следующим кодом:

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. Сохраните файл.

## <a name="code-the-server-side"></a>Код на стороне сервера

### <a name="configure-the-owin-middleware"></a>Настройка ПО промежуточного слоя OWIN

1. Откройте файл Startup.cs в корневой папке проекта **Office-Add-in-ASPNET-SSO-WebAPI** и добавьте приведенный ниже метод в класс **Startup**. Обратите внимание, что метод `ConfigureAuth` создается позже.

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. Сохраните и закройте файл.

1. Щелкните правой кнопкой мыши папку **App_Start** и выберите **Добавить > Класс**.

1. В диалоговом окне **Добавить новый элемент** введите имя файла **Startup.Auth.cs** и нажмите кнопку **Добавить**.

1. Сократите имя пространства имен в новом файле до `Office_Add_in_ASPNET_SSO_WebAPI`.

1. Убедитесь, что в начале файла есть все приведенные ниже операторы `using`.

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:

    `public partial class Startup`

1. Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. Замените `TODO 1` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Код указывает OWIN, чтобы убедиться, что аудитория, указанная в маркере начальной загрузки, который поступает из приложения Office, должна соответствовать значению, указанному в web.config.
    - У учетных записей Майкрософт есть идентификатор GUID издателя, отличный от GUID любого клиента организации, поэтому для поддержки обоих типов учетных записей мы не проверяем издателя.
    - Если задано значение `SaveSigninToken` , `true` OWIN будет сохранять необработанный маркер начальной загрузки из приложения Office. Он необходим надстройке, чтобы получить маркер доступа к Microsoft Graph в потоке "от имени".
    - ПО промежуточного слоя OWIN не проверяет области. Области маркера начальной загрузки, которые должны включать `access_as_user`, проверяются в контроллере.

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. Замените `TODO 2` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Метод `UseOAuthBearerAuthentication` вызывается вместо более распространенного метода `UseWindowsAzureActiveDirectoryBearerAuthentication`, так как последний несовместим с конечной точкой Azure AD версии 2.
    - URL-адрес, передаваемый методу, — это то, где ПО промежуточного слоя OWIN получает инструкции по получению ключа, необходимого для проверки подписи маркера начальной загрузки, полученного от приложения Office. Сегмент URL-адреса "Полномочия" предоставляется файлом web.config. Это либо строка "common", либо GUID для однотенантной надстройки.

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. Сохраните и закройте файл.

### <a name="create-the-apivalues-controller"></a>Создание контроллера /api/values

1. Откройте файл **Controllers\ValueController.cs**. Этот контроллер используется в случае успешного получения маркера начальной загрузки системой единого входа. Он не используется в рамках резервной системы авторизации. В этой системе использован AzureADAuthController, созданный для вас.

1. Убедитесь, что в начале файла есть приведенные ниже инструкции с `using`.

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

1. Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.

1. Добавьте приведенный ниже метод в `ValuesController`. Обратите внимание, что возвращаемое значение — `Task<HttpResponseMessage>`, а не `Task<IEnumerable<string>>`, которое чаще используется для метода `GET api/values`. Это является побочным эффектом того факта, что логика авторизации OAuth должна находиться в контроллере, а не в фильтре ASP.NET. Некоторые условия возникновения ошибки в этой логике требуют отправки объекта HTTP-ответа в клиент надстройки.

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //         token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get a new access token for Microsoft Graph.

        // TODO 4: Use the new access token to call Microsoft Graph.
    }
    ```

1. Замените `TODO1` приведенным ниже кодом, чтобы убедиться, что в маркере указано разрешение `access_as_user`. Обратите внимание, что второй параметр метода `SendErrorToClient` — объект **Exception**. В этом случае код передает `null`, потому что включение объекта **Exception** блокирует включение свойства **Message** в создаваемый HTTP-ответ.

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. Замените `TODO 2` приведенным ниже кодом, чтобы собрать все сведения, необходимые для получения маркера для Microsoft Graph, используя поток "от имени". Вот что нужно знать об этом коде:

    - Ваша надстройка больше не играет роль ресурса (или аудитории), к которым приложению Office и пользователю требуется доступ. Теперь она сама является клиентом, которому необходим доступ к Microsoft Graph. `ConfidentialClientApplication` — это объект "контекста клиента" MSAL.
    - Начиная с MSAL.NET 3.x.x, `bootstrapContext` — это сам маркер начальной загрузки. 
    - Полномочия предоставляются файлом web.config. Это либо строка "common", либо GUID для однотенантной надстройки.
    - MSAL выдает ошибку, если код запрашивает `profile`, который действительно используется только в том случае, если клиентское приложение Office получает маркер веб-приложению надстройки. Поэтому явным образом запрашивается только `Files.Read.All`.

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri(ConfigurationManager.AppSettings["ida:Domain"])
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. Замените `TODO 3` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Для начала метод `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` проверит кэш MSAL, который находится в памяти, на наличие подходящего маркера доступа. Только в случае его отсутствия запускается поток "от имени" с конечной точкой Azure AD версии 2.
    - Любые исключения, отличные от типа `MsalServiceException`, не перехватываются преднамеренно, поэтому будут переданы клиенту в виде сообщений `500 Server Error`.

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

1. Замените `TODO 3a` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Если ресурс Microsoft Graph требует многофакторной проверки подлинности, а пользователь еще не предоставил соответствующие данные, Azure AD вернет состояние "400 Bad Request" с ошибкой `AADSTS50076` и свойство **Claims**. MSAL выдает исключение **MsalUiRequiredException** (которое наследуется от **MsalServiceException**), используя эту информацию.
    - Значение свойства **Claims** должно быть передано клиенту, который должен передать его в приложение Office, которое затем включает его в запрос на новый маркер начальной загрузки. Azure AD предложит пользователю пройти все необходимые проверки подлинности.
    - The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.
    - Сообщение создается в формате JSON, чтобы клиентский код JavaScript мог проанализировать его с помощью известных методов объекта JavaScript `JSON`.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. Замените `TODO 3b` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Если вызов Azure AD содержал по крайней мере одно разрешение, которое не предоставил ни пользователь, ни администратор клиента (или оно было отозвано), Azure AD вернет состояние "400 Bad Request" с ошибкой `AADSTS65001`. MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию.
    - Если вызов Azure AD содержал по крайней мере одно нераспознанное разрешение, Azure AD вернет состояние "400 Bad Request" с ошибкой `AADSTS70011`. MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию.
    - Полное описание включается, так как ошибка 70011 возвращается и в других случаях, и ее следует обрабатывать в этой надстройке, только когда она означает запрос недопустимого разрешения.
    - The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. Замените `TODO 3c` приведенным ниже кодом, чтобы обработать все остальные исключения **MsalServiceException**.

    ```csharp
    else
    {
        throw e;
    }
    ```

1. замените `TODO 4` приведенным ниже кодом. Метод `GraphApiHelper.GetOneDriveFileNames`, созданный для вас, выполняет запрос данных в Microsoft Graph и включает маркер доступа.

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. Сохраните и закройте файл.

## <a name="run-the-solution"></a>Запуск решения

1. Откройте файл решения в Visual Studio.
1. В меню **Построение** выберите команду **Очистить решение**. После выполнения команды снова откройте меню **Построение** и выберите команду **Построить решение**.
1. В **обозревателе решений** выберите узел проекта **Office-Add-in-ASPNET-SSO** (не верхний узел решения и не узел проекта, имя которого заканчивается на "WebAPI").
1. В области **Свойства** откройте раскрывающийся список **Начальный документ** и выберите один из трех вариантов (Excel, Word или PowerPoint).

    ![Выберите нужное клиентское приложение Office: Excel, PowerPoint или Word.](../images/SelectHost.JPG)

1. Нажмите клавишу F5.
1. В приложении Office на вкладке ленты **Главная** в группе **Единый вход ASP.NET** выберите команду **Показать надстройку**, чтобы открыть надстройку области задач.
1. Нажмите кнопку **Получить имена файлов OneDrive**. Если вы вошли в Office с помощью Microsoft 365 для образования, рабочей учетной записи или учетной записи Майкрософт и единый вход работает должным образом, первые 10 имен файлов и папок в OneDrive для бизнеса отображаются в области задач. Если вы не вошли в систему или находитесь в сценарии, который не поддерживает единый вход, или единый вход не работает по какой-либо причине, вам будет предложено выполнить вход. После входа отображаются имена файлов и папок.

### <a name="testing-the-fallback-path"></a>Тестирование резервного пути

Чтобы протестировать резервный путь авторизации, принудим путь единого входа к ошибке, выполнив следующие действия.

1. Добавьте следующий код в самую верхнюю часть `getDataWithToken` метода в файле HomeES6.js.

    ```javascript
    function MockSSOError(code) {
        this.code = code;
    }
    ```

1. Затем добавьте следующую строку в верхнюю часть `try` блока в том же методе сразу над вызовом `getAccessToken`.

    ```javascript
    throw new MockSSOError("13003");
    ```

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a>Обновление надстройки при переходе к промежуточной и рабочей среде

Как и все веб-надстройки Office, когда вы готовы перейти на промежуточный или рабочий сервер, необходимо обновить `localhost:44355` домен в манифесте новым доменом. Аналогичным образом необходимо обновить домен в файле web.config.

Так как домен отображается в регистрации AAD, необходимо обновить эту регистрацию, чтобы использовать новый домен вместо `localhost:44355` того, где бы он ни появился.
