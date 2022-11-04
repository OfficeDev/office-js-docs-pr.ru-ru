---
title: Создание надстройки Office на платформе Node.js с использованием единого входа
description: Узнайте, как создать надстройку на основе Node.js, которая использует единый вход Office.
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 35128da43b3f27a58df5e188a5001bfa8aba4a4c
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/28/2022
ms.locfileid: "68841789"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>Создание надстройки Office на платформе Node.js с использованием единого входа

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

В этой статье описывается процесс включения единого входа (SSO) в надстройке. Создаваемый образец надстройки состоит из двух частей. область задач, загружаемая в Microsoft Excel, и сервер среднего уровня, обрабатывающий вызовы Microsoft Graph для области задач. Сервер среднего уровня создается с помощью Node.js и Express и предоставляет один REST API , `/getuserfilenames`который возвращает список первых 10 имен файлов в папке OneDrive пользователя. Область задач использует `getAccessToken()` метод для получения маркера доступа для пользователя, выполнившего вход на сервер среднего уровня. Сервер среднего уровня использует поток On-Behalf-Of (OBO) для обмена маркером доступа на новый с доступом к Microsoft Graph. Этот шаблон можно расширить для доступа к любым данным Microsoft Graph. Область задач всегда вызывает REST API среднего уровня (передавая маркер доступа), когда ей требуются службы Microsoft Graph. Средний уровень использует маркер, полученный через OBO, для вызова служб Microsoft Graph и возврата результатов в область задач.

Эта статья работает с надстройкой, которая использует Node.js и Express. Аналогичная статья, посвященная надстройке на основе ASP.NET, — [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Необходимые компоненты

- [Node.js](https://nodejs.org/) (последняя версия [LTS](https://nodejs.org/about/releases))

- [Git Bash](https://git-scm.com/downloads) (или другой клиент git).

- Редактор кода— мы рекомендуем Visual Studio Code

- По крайней мере несколько файлов и папок, хранящихся в OneDrive для бизнеса в подписке Microsoft 365

- Сборка Microsoft 365, поддерживающая [набор требований IdentityAPI 1.3](/javascript/api/requirement-sets/common/identity-api-requirement-sets). Вы можете получить [бесплатную песочницу для разработчиков](https://developer.microsoft.com/microsoft-365/dev-program#Subscription), которая предоставляет возобновляемую 90-дневную подписку Microsoft 365 E5 разработчика. Песочница разработчика включает подписку Microsoft Azure, которую можно использовать для регистрации приложений на последующих шагах в этой статье. При желании для регистрации приложений можно использовать отдельную подписку Microsoft Azure. Получите пробную подписку на [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Настройка начального проекта

1. Клонируйте или скачайте репозиторий [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO).

   > [!NOTE]
   > Существует две версии примера.
   >
   > - Папка **Begin** является начальным проектом. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы. В последующих разделах этой статьи рассматривается доработка проекта.
   > - Папка **Complete** содержит один и тот же пример со всеми инструкциями по написанию кода из этой статьи. Чтобы использовать завершенную версию, просто следуйте инструкциям в этой статье, но замените "Begin" на "Complete" (Завершить) и пропустите разделы **Код на стороне клиента** и **Код на стороне сервера среднего уровня** .

1. Откройте командную строку в папке **Begin** .

1. Введите в консоли команду `npm install`, чтобы установить все зависимости, указанные в файле package.json.

1. Выполните команду `npm run install-dev-certs`. При запросе нажмите **Да**, чтобы установить сертификат.

Используйте следующие значения заполнителей для последующих шагов регистрации приложения.

| Заполнитель           | Значение                                 |
|-----------------------|---------------------------------------|
| `<add-in-name>`       | **Office-Add-in-NodeJS-SSO**          |
| `<redirect-platform>` | **Одностраничное приложение (SPA)**     |
| `<redirect-uri>`      | `https://localhost:44355/dialog.html` |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="configure-the-add-in"></a>Настройка надстройки

1. Откройте папку `\Begin` в скопированном проекте в редакторе кода.

1. `.ENV` Откройте файл и используйте значения, скопированные ранее из регистрации приложения **Office-Add-in-NodeJS-SSO**. Задайте значения следующим образом:

   | Имя              | Значение                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | **Идентификатор приложения (клиента)** на странице обзора регистрации приложения. |
   | **CLIENT_SECRET** | **Секрет клиента** , **сохраненный на странице "Сертификаты & секреты** ".       |
   | **DIRECTORY_ID**  | **Идентификатор каталога (клиента)** на странице обзора регистрации приложения.   |

   Значения **не** должны быть заключены в кавычки. По завершении файл должен выглядеть следующим образом.

   ```javascript
   CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
   CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
   DIRECTORY_ID=478aa78e-20ba-4c0d-9ffe-c4f62e5de3d5
   NODE_ENV=development
   SERVER_SOURCE=https://localhost:44355

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup.

   ```xml
   <WebApplicationInfo>
     <Id>$app-id-guid$</Id>
     <Resource>api://localhost:44355/$app-id-guid$</Resource>
     <Scopes>
         <Scope>Files.Read</Scope>
         <Scope>profile</Scope>
         <Scope>openid</Scope>
     </Scopes>
   </WebApplicationInfo>
   ```

1. Замените заполнитель "$app-id-guid$" _в обоих местах_ разметки **идентификатором приложения** , скопированным при создании регистрации приложения **Office-Add-in-NodeJS-SSO** . Символы "$" не являются частью идентификатора, поэтому не включайте их. Это тот же идентификатор, который использовался для CLIENT_ID в . ENV-файл.

   > [!NOTE]
   > Значением **\<Resource\>** является **URI идентификатора приложения** , заданный при регистрации надстройки. Раздел **\<Scopes\>** используется только для создания диалогового окна согласия, если надстройка продается через AppSource.

1. Откройте файл `\public\javascripts\fallback-msal\authConfig.js`. Замените заполнитель "$app-id-guid$" идентификатором приложения, сохраненным из созданной ранее регистрации приложения **Office-Add-in-NodeJS-SSO** .

1. Сохраните изменения в файле.

## <a name="code-the-client-side"></a>Код на стороне клиента

### <a name="create-client-request-and-response-handler"></a>Создание обработчика запросов и ответов клиента

1. Откройте файл `public\javascripts\ssoAuthES6.js` в редакторе кода. В нем уже есть код, обеспечивающий поддержку обещаний (даже в Internet Explorer 11), и вызов `Office.onReady` для назначения обработчика единственной кнопки надстройки.

   > [!NOTE]
   > Как следует из названия, ssoAuthES6.js использует синтаксис JavaScript ES6, так как применение `async` и `await` хорошо демонстрирует простоту API единого входа. При запуске сервера localhost этот файл преобразуется в синтаксис ES5, чтобы пример поддерживал Internet Explorer 11.

    Ключевой частью примера кода является клиентский запрос. Клиентский запрос — это объект, отслеживающий сведения о запросе для вызова REST API на сервере среднего уровня. Это необходимо, так как состояние запроса клиента необходимо отслеживать или обновлять с помощью следующего сценария:

    - Единый вход завершается сбоем, и требуется резервная проверка подлинности. Маркер доступа получается через MSAL во всплывающем диалоговом окне. Цель заключается в том, чтобы не завершиться неудачей в этом сценарии и корректно вернуться к альтернативному подходу проверки подлинности.

    Объект запроса клиента отслеживает следующие данные:

    - `authSSO` — значение true, если используется единый вход, в противном случае — значение false.
    - `verb` — команды REST API, такие как GET и POST.
    - `accessToken`— маркер доступа к серверу ASP.NET Core.
    - `url`— URL-адрес REST API для вызова на сервере ASP.NET Core.
    - `callbackRESTApiHandler` — Функция для передачи результатов вызова REST API.
    - `callbackFunction` — функция, передаваемая клиентский запрос при готовности.

1. Чтобы инициализировать объект клиентского запроса, замените `createRequest` в функции `TODO 1` следующим кодом.

    ```javascript
    const clientRequest = {
      authSSO: authSSO,
      verb: verb,
      accessToken: null,
      url: url,
      callbackRESTApiHandler: restApiCallback,
        callbackFunction: callbackFunction,
    };
    ```

1. Замените `TODO 2` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Он проверяет, используется ли единый вход. Метод получения маркера доступа для единого входа отличается от метода резервной проверки подлинности.
    - Если единый вход возвращает маркер доступа, он вызывает функцию `callbackfunction` . Для резервной проверки подлинности он вызывает `dialogFallback`, который в конечном итоге вызывает функцию обратного вызова после входа пользователя через MSAL.

    ```javascript
    // Get access token.

    if (authSSO) {
    try {
      // Get access token from Office SSO.
      clientRequest.accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
      callbackFunction(clientRequest);
    } catch (error) {
      // handle the SSO error which will inform us if we need to switch to fallback auth.
      let fallbackRequired = handleSSOErrors(error);
      if (fallbackRequired) switchToFallbackAuth(clientRequest);
    }
   } else {
     // Use fallback auth to get access token.
     dialogFallback(clientRequest);
   }
    ```

1. В функции `getFileNameList` замените `TODO 3` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Функция `getFileNameList` вызывается, когда пользователь нажимает кнопку **Получить имена файлов OneDrive** в области задач.
    - Он создает клиентский запрос для отслеживания сведений о вызове, таких как URL-адрес REST API.
    - Когда REST API возвращает результат, он передается в функцию `handleGetFileNameResponse` . Этот обратный вызов передается в качестве параметра `createRequest` и отслеживается в `clientRequest.callbackRESTApiHandler`.
    - Код вызывает `callWebServer` с клиентским запросом для выполнения дальнейших действий и вызова REST API.

    ```javascript
    createRequest(
      "GET",
      "/getuserfilenames",
      handleGetFileNameResponse,
      async (clientRequest) => {
        await callWebServer(clientRequest);
      }
    );
    ```

1. В функции `handleGetFileNameResponse` замените `TODO 4` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Код передает ответ (который содержит список имен файлов) для `writeFileNamesToOfficeDocument` записи имен файлов в документ.
    - Код проверяет наличие ошибок. В нем отображается сообщение об успешном выполнении, если имена файлов записаны, в противном случае отображается ошибка.

    ```javascript
    if (response !== null) {
      try {
        await writeFileNamesToOfficeDocument(response);
        showMessage("Your OneDrive filenames are added to the document.");
      } catch (error) {
        // The error from writeFileNamesToOfficeDocument will begin
        // "Unable to add filenames to document."
        showMessage(error);
      }
    } else
    showMessage("A null response was returned to handleGetFileNameResponse.");
    ```

1. В функции `handleSSOErrors` замените `TODO 5` приведенным ниже кодом. Дополнительные сведения об этих ошибках см. в статье [Устранение ошибок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md).

    ```javascript
    let fallbackRequired = false;

   switch (err.code) {
     case 13001:
       // No one is signed into Office. If the add-in cannot be effectively used when no one
       // is logged into Office, then the first call of getAccessToken should pass the
       // `allowSignInPrompt: true` option. Since this sample does that, you should not see
       // this error.
       showMessage(
         "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
       );
       break;
     case 13002:
       // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
       // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
       showMessage(
         "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
       );
       break;
     case 13006:
       // Only seen in Office on the web.
       showMessage(
         "Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
       );
       break;
     case 13008:
       // Only seen in Office on the web.
       showMessage(
        "Office is still working on the last operation. When it completes, try this operation again."
       );
       break;
     case 13010:
       // Only seen in Office on the web.
       showMessage(
         "Follow the instructions to change your browser's zone configuration."
       );
       break;
    ```

1. Замените `TODO 6` приведенным ниже кодом. Дополнительные сведения об этих ошибках см. [в статье Устранение неполадок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md). Для любых ошибок, которые не могут быть обработаны, `true` возвращается вызывающей. Это означает, что вызывающий объект должен переключиться на использование MSAL в качестве резервной проверки подлинности.

    ```javascript
     default:
      // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
      // to non-SSO sign-in.
      fallbackRequired = true;
      break;
    }
    return fallbackRequired;
    ```

### <a name="call-the-rest-api-on-the-middle-tier-server"></a>Вызов REST API на сервере среднего уровня

1. В функции `callWebServer` замените `TODO 7` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Фактический вызов AJAX будет выполнен функцией `ajaxCallToRESTApi` .
    - Эта функция попытается получить новый маркер доступа, если сервер среднего уровня возвращает ошибку, указывающую, что срок действия текущего маркера истек.
    - Если вызов AJAX не может быть выполнен успешно, `switchToFallbackAuth` будет вызван для использования проверки подлинности MSAL вместо единого входа Office.

    ```javascript
    try {
    const data = await $.ajax({
      type: clientRequest.verb,
      url: clientRequest.url,
      headers: { Authorization: "Bearer " + clientRequest.accessToken },
      cache: false,
    });
    clientRequest.callbackRESTApiHandler(data);

    } catch (error) {
     // TODO 8: Check for expired SSO token and refresh if needed.

    // TODO 9: Check for Microsoft Graph and other errors.

    }
    ```

1. Замените `TODO 8` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Когда сервер определяет маркер с истекшим сроком действия, он возвращает ошибку с типом TokenExpiredError.
    - Попробуйте... catch вызовет Office.auth.getAccessToken, чтобы получить обновленный маркер с новым сроком действия.
    - Код попытается снова вызвать API сервера.

    ```javascript
    // Check for expired SSO token. Refresh and retry the call if it expired.
    if (
      error.responseJSON &&
      authSSO === true &&
      error.responseJSON.type === "TokenExpiredError"
    ) {
      try {
        const accessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
        const data = await $.ajax({
          type: clientRequest.verb,
          url: clientRequest.url,
          headers: { Authorization: "Bearer " + accessToken },
          cache: false,
        });
        clientRequest.callbackRESTApiHandler(data);
      } catch (error) {
        showMessage(error.responseText);
        switchToFallbackAuth(clientRequest);
        return;
      }
    }
    ```

1. Замените `TODO 9` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Для ошибок **Microsoft Graph** покажите сообщение в области задач.
    - Для всех остальных сообщений покажите сообщение в области задач.

    ```javascript
    // Check for a Microsoft Graph API call error. which is returned as bad request (403)
    if (error.status === 403) {
      if (error.responseJSON && error.responseJSON.type === "Microsoft Graph") {
        showMessage(error.responseJSON.errorDetails);
      } else {
        showMessage(error);
      }
      return;
    }

    // For all other error scenarios, display the message and use fallback auth.
    showMessage("Unknown error from web server: " + JSON.stringify(error));
    if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
    ```

Резервная проверка подлинности использует библиотеку MSAL для входа пользователя. Сама надстройка является SPA и использует регистрацию приложения SPA для доступа к серверу среднего уровня.

1. В функции `switchToFallbackAuth` замените `TODO 10` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Он устанавливает для глобального `authSSO` значения значение false и создает новый клиентский запрос, который использует MSAL для проверки подлинности. Новый запрос содержит маркер доступа MSAL к серверу среднего уровня.
    - После создания запроса он вызывает `callWebServer` для продолжения попытки успешного вызова сервера среднего уровня.

    ```javascript
    // Guard against accidental call to this function when fallback is already in use.

    if (authSSO === false) return;

    showMessage("Switching from SSO to fallback auth.");
    authSSO = false;
    // Create a new request for fallback auth.
    createRequest(
      clientRequest.verb,
      clientRequest.url,
      clientRequest.callbackRESTApiHandler,
      async (fallbackRequest) => {
        // Hand off to call using fallback auth.
        await callWebServer(fallbackRequest);
      }
    );
    ```

## <a name="code-the-middle-tier-server"></a>Код сервера среднего уровня

Сервер среднего уровня предоставляет REST API для вызова клиентом. Например, REST API `/getuserfilenames` получает список имен файлов из папки OneDrive пользователя. Каждому вызову REST API требуется маркер доступа клиента, чтобы убедиться, что правильный клиент обращается к своим данным. Маркер доступа обменивается на маркер Microsoft Graph через поток On-Behalf-Of (OBO). Новый токен Microsoft Graph кэшируется библиотекой MSAL для последующих вызовов API. Он никогда не отправляется за пределы сервера среднего уровня. Дополнительные сведения см. в разделе [Запрос маркера доступа среднего уровня](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request).

### <a name="create-the-route-and-implement-on-behalf-of-flow"></a>Создание маршрута и реализация потока On-Behalf-Of

1. Откройте файл `routes\getFilesRoute.js` и замените `TODO 11` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Он вызывает .`authHelper.validateJwt` Это гарантирует, что маркер доступа действителен и не был изменен.
    - Дополнительные сведения см. в разделе [Проверка маркеров](/azure/active-directory/develop/access-tokens#validating-tokens).

    ```javascript
    router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 12: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
    );
    ```

1. Замените `TODO 12` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Он запрашивает только необходимые минимальные области, например `files.read`.
    - Он использует MSAL `authHelper` для выполнения потока OBO в вызове `acquireTokenOnBehalfOf`.

    ```javascript
    try {
      const authHeader = req.headers.authorization;
      let oboRequest = {
        oboAssertion: authHeader.split(" ")[1],
        scopes: ["files.read"],
      };

      // The Scope claim tells you what permissions the client application has in the service.
      // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
      const accessAsUserScope = tokenScopes.find(
        (scope) => scope === "access_as_user"
      );
      if (!accessAsUserScope) {
        res.status(401).send({ type: "Missing access_as_user" });
        return;
      }
      const cca = authHelper.getConfidentialClientApplication();
      const response = await cca.acquireTokenOnBehalfOf(oboRequest);
      // TODO 13: Call Microsoft Graph to get list of filenames.
    } catch (err) {
      // TODO 14: Handle any errors.
    }
    ```

1. Замените `TODO 13` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Он создает URL-адрес для вызова Microsoft API Graph, а затем выполняет вызов через функцию `getGraphData` .
    - Он возвращает ошибки, отправляя ответ HTTP 500 вместе с подробными сведениями.
    - При успешном выполнении он возвращает клиенту json со списком имен файлов.

    ```javascript
    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 10 folder or file names.
    const rootUrl = "/me/drive/root/children";

    // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
    // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
    // sanitized so that it cannot be used in a Response header injection attack.
    const params = "?$select=name&$top=10";

    const graphData = await getGraphData(
      response.accessToken,
      rootUrl,
      params
    );

    // If Microsoft Graph returns an error, such as invalid or expired token,
    // there will be a code property in the returned object set to a HTTP status (e.g. 401).
    // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
    if (graphData.code) {
      res
        .status(403)
        .send({
          type: "Microsoft Graph",
          errorDetails:
            "An error occurred while calling the Microsoft Graph API.\n" +
            graphData,
        });
    } else {
      // MS Graph data includes OData metadata and eTags that we don't need.
      // Send only what is actually needed to the client: the item names.
      const itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
      }

      res.status(200).send(itemNames);
    }
    ```

1. Замените `TODO 14` на приведенный ниже код. Этот код специально проверяет, истек ли срок действия маркера, так как клиент может запросить новый маркер и снова вызвать.

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(401).send({ type: "TokenExpiredError", errorDetails: err });
   } else {
     res.status(403).send({ type: "Unknown", errorDetails: err });
   }
   ```

Пример должен обрабатывать как резервную проверку подлинности через MSAL, так и проверку подлинности единого входа через Office. Пример сначала попытается выполнить единый вход, а `authSSO` логическое значение в верхней части файла будет отслеживаться, если пример использует единый вход или переключился на резервную проверку подлинности.

## <a name="run-the-project"></a>Запуск проекта

1. Убедитесь в наличии нескольких файлов в OneDrive, чтобы можно было проверить результаты.

1. Откройте командную строку в корне папки `\Begin`.

1. Выполните команду `npm install` , чтобы установить все зависимости пакета.

1. Выполните команду `npm start` , чтобы запустить сервер среднего уровня.

1. Вам потребуется загрузить неопубликованную надстройку в приложение Office (Excel, Word или PowerPoint), чтобы протестировать ее. Инструкции зависят от вашей платформы. Ссылки на инструкции доступны в разделе [Загрузка неопубликованной надстройки Office для тестирования](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. В приложении Office на вкладке ленты **Главная** нажмите кнопку **Показать надстройку** в группе **Единый вход Node.js**, чтобы открыть надстройку области задач.

1. Нажмите кнопку **Получить имена файлов OneDrive**. Если вы вошли в Office с помощью Microsoft 365 для образования или рабочей учетной записи или учетной записи Майкрософт и единый вход работает должным образом, первые 10 имен файлов и папок в вашем OneDrive для бизнеса вставляются в документ. (Первый раз может занять до 15 секунд.) Если вы не вошли в систему или находитесь в сценарии, который не поддерживает единый вход, или единый вход не работает по какой-либо причине, вам будет предложено войти. После входа отображаются имена файлов и папок.

> [!NOTE]
> Если вы ранее выполняли вход в Office с использованием другого идентификатора и все еще не закрыли некоторые из открытых тогда приложений Office, Office может не сменить идентификатор (даже если кажется, что это сделано). Если это произойдет, возможен сбой при вызове Microsoft Graph или возврат данных для другого идентификатора. Чтобы избежать этого, _закройте все приложения Office_, прежде чем нажимать кнопку **Получить имена файлов OneDrive**.

## <a name="security-notes"></a>Заметки о безопасности

- Маршрут `/getuserfilenames` в `getFilesroute.js` использует литеральную строку для создания вызова Microsoft Graph. Если вы измените вызов так, чтобы какая-либо часть строки попадала из входных данных пользователя, очищайте входные данные, чтобы их нельзя было использовать при атаке с внедрением заголовка ответа.

- В `app.js` приведенной ниже политике безопасности содержимого применяется для сценариев. Вы можете указать дополнительные ограничения в зависимости от потребностей надстройки в безопасности.

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

Всегда следуйте рекомендациям по обеспечению безопасности в [документации по платформа удостоверений Майкрософт](/azure/active-directory/develop/).
