---
title: Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET
description: Пошаговое руководство по созданию (или преобразованию) надстройки Office с серверной ASP.NET для использования единого входа (SSO).
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 403730f953a4f53d853a0ecd3b12cd477f7e7176
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958834"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET

После того как пользователи войдут в Office, ваша надстройка сможет использовать те же учетные данные для предоставления им доступа к нескольким приложениям без необходимости повторного входа. Общие сведения см. в статье [Включение единого входа в надстройке Office](sso-in-office-add-ins.md).
В этой статье описывается процесс включения единого входа (SSO) в надстройке, созданной с ASP.NET.

## <a name="prerequisites"></a>Предварительные требования

- Visual Studio 2019 или более поздней версии.

- Рабочая **нагрузка разработки Office или SharePoint** при настройке Visual Studio.

- [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

- По крайней мере несколько файлов и папок, хранящихся OneDrive для бизнеса в подписке Microsoft 365.

- Учетная запись Azure с активной подпиской [— создайте учетную запись бесплатно](https://azure.microsoft.com/free/?WT.mc_id=A261C142F).

## <a name="set-up-the-starter-project"></a>Настройка начального проекта

Клонируйте или скачайте репозиторий [Office Add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).

> [!NOTE]
> Существует две версии примера.
>
> - В папке **Before** находится начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы. В последующих разделах этой статьи рассматривается доработка проекта.
> - Версия примера в папке **Complete** идентична надстройке, которую вы бы создали, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью. Чтобы использовать готовую версию, просто выполните действия, описанные в этой статье, но замените папку "Before" на папку "Complete" и пропустите разделы **Код на стороне клиента** и **Код на стороне сервера**.

## <a name="register-the-add-in-through-an-app-registration"></a>Регистрация надстройки с помощью регистрации приложения

Сначала выполните действия, описанные в кратком руководстве[.](/azure/active-directory/develop/quickstart-register-app) Регистрация приложения платформа удостоверений Майкрософт регистрации надстройки.

Используйте следующие параметры для регистрации приложения.

- Имя: `Office-Add-in-ASPNET-SSO`
- Поддерживаемые типы учетных записей: учетные записи в любом каталоге организации (любой каталог Azure AD — мультитенантный) и личные учетные записи **Майкрософт (например, Skype, Xbox)**

    > [!NOTE]
    >  Если вы хотите, чтобы надстройка была доступна только пользователям в клиенте, где вы ее регистрируете, вы можете выбрать учетные записи только в этом каталоге организации, но вам потребуется выполнить некоторые дополнительные действия по настройке. **Дополнительные сведения о настройке для одного клиента** см. далее в этой статье.

- Платформа: **Интернет**
- URI перенаправления: **https://localhost:44355/AzureADAuth/Authorize**
- Секрет клиента: `*********` (веб-приложение использует секрет клиента для подтверждения своей личности при запросе маркеров. *Запишите это значение для использования на следующем шаге — оно отображается только один раз.*)

### <a name="expose-a-web-api"></a>Предоставление веб-API

1. В созданной регистрации приложения выберите **"Предоставить API> добавить область**.
   Вам будет предложено задать универсальный код ресурса **(URI) идентификатора** приложения, если он еще не настроен.

    URI идентификатора приложения выступает в качестве префикса для областей, на которые вы будете ссылаться в коде API, и должен быть глобально уникальным. Используйте форму, `api://localhost:44355/[application-id-guid]`например `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. Укажите атрибуты области **в области добавления** области.

    |Поле          |Значение  |
    |---------------|---------|
    |**Имя области** | `access_as_user`|
    |**Кто может дать согласие** | **Администраторы и пользователи**|
    |**Администратор отображаемое имя согласия** | Office может выступать в качестве пользователя.|
    |**Администратор согласия** | Разрешить Office вызывать веб-API надстройки с правами текущего пользователя.|
    |**Отображаемое имя согласия пользователя** | Office может действовать от имени вас.|
    |**Описание согласия пользователя** | Разрешите Office вызывать веб-API надстройки с правами, которые у вас есть.|

1. **Задайте для параметра "** Состояние **" значение "Включено**", а затем выберите "**Добавить область"**.

    > [!NOTE]
    > Доменная часть **\<Scope\>** имени, отображаемая сразу после текстового поля, должна автоматически соответствовать заданным ранее URI идентификатора приложения с `/access_as_user` добавлением к концу, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`например .

1. В разделе **"Авторизованные** клиентские приложения" введите следующий идентификатор, чтобы предварительно авторизовать все конечные точки приложений Microsoft Office.

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Все конечные точки приложения Microsoft Office)

    > [!NOTE]
    > Идентификатор `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` предварительно авторизует Office на всех следующих платформах. Кроме того, можно ввести соответствующее подмножество следующих идентификаторов, если по какой-либо причине вы хотите запретить авторизацию в Office на некоторых платформах. Просто оставьте идентификаторы платформ, с которых требуется отостановить авторизацию. Пользователи надстройки на этих платформах не смогут вызывать веб-API, но другие функции надстройки по-прежнему будут работать.
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).
    > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office в Интернете).
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook в Интернете).

1. Выберите **Добавить клиентское приложение**. На открываемой панели задайте для идентификатора клиента соответствующий GUID и установите флажок для `api://localhost:44355/[application-id-guid]/access_as_user`.

1. Нажмите кнопку **Добавить приложение**.

### <a name="configure-microsoft-graph-permissions"></a>Настройка разрешений Microsoft Graph

1. Выберите **разрешения API > добавить разрешение > Microsoft Graph**.

1. Выберите **Делегированные разрешения**. Microsoft Graph предоставляет множество разрешений, наиболее часто используемых в верхней части списка.

1. В **разделе "Выбор разрешений**" выберите следующие разрешения.

    |Разрешение     |Описание  |
    |---------------|-------------|
    |Files.Read.All |Чтение всех файлов, к которых пользователь может получить доступ. |
    |profile        |Просмотр базового профиля пользователей. Требуется приложению Office для получения маркера для веб-приложения надстройки. |

    > [!NOTE]
    > Разрешение `User.Read` может быть уже указано по умолчанию. Незачем запрашивать ненужные разрешения, поэтому рекомендуем снять флажок рядом с разрешением, которое не требуется вашей надстройке.

1. Выберите **"Добавить разрешения"** , чтобы завершить процесс.

Каждый раз, когда вы настраиваете разрешения, пользователи приложения запрашиваются при входе в систему, чтобы предоставить приложению доступ к API ресурсов от их имени. Администратор также может предоставить согласие от имени всех пользователей, чтобы им не было предложено сделать это.

1. На этой же странице нажмите кнопку **Предоставить согласие администратора для [имя клиента]** и выберите **Принять** в появившемся запросе подтверждения.

    > [!NOTE]
    > После нажатия кнопки **Предоставить согласие администратора для [имя клиента]** может появиться сообщение баннера с просьбой повторить попытку через несколько минут, чтобы можно было создать запрос на продолжение. В этом случае можно приступить к работе со следующим разделом, но не забудьте вернуться на **_портал и нажать эту кнопку_**.

## <a name="configure-the-solution"></a>Настройка решения

1. В корне папки **Before** откройте SLN-файл решения в **Visual Studio**. В **обозревателе решений** щелкните правой кнопкой мыши верхний узел (узел решения, а не узлы проектов) и выберите **Назначить запускаемые проекты**.

1. В разделе **Общие свойства** выберите **Запускаемый проект**, а затем **Несколько запускаемых проектов**. Убедитесь, что для параметра **Действие** в обоих проектах установлено значение **Запуск** и что проект, заканчивающийся на "...WebAPI", указан в списке первым. Закройте диалоговое окно.

1. Вернитесь **Обозреватель решений** выберите (не щелкайте правой кнопкой мыши) проект **Office-Add-in-ASPNET-SSO-WebAPI**. Откроется область **Свойства**. Убедитесь, что для параметра **SSL включен** задано значение **True**. Убедитесь, что **URL-адрес SSL** указан как `http://localhost:44355/`.

1. В файле web.config используйте значения, скопированные ранее. Для **ida:ClientID** и **ida:Audience** укажите **идентификатор приложения (клиента)**, для **ida:Password** — секрет клиента. Кроме того, **задайте для ida:Domain** значение `http://localhost:44355` (без косой черты "/" в конце).

    > [!NOTE]
    > Идентификатор **приложения (клиента)** — это значение аудитории, когда другие приложения, такие как клиентское приложение Office (например, PowerPoint, Word, Excel), ищут авторизованный доступ к приложению. Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.

1. Если вы не указали вариант "Учетные записи только в этом каталоге организации" для параметра **ПОДДЕРЖИВАЕМЫЕ ТИПЫ УЧЕТНЫХ ЗАПИСЕЙ** при регистрации настройки, сохраните и закройте файл web.config. В противном случае сохраните его, но оставьте открытым. 

1. В **Обозреватель решений выберите** проект **Office-Add-in-ASPNET-SSO** и откройте файл манифеста надстройки "Office-Add-in-ASPNET-SSO.xml", а затем прокрутите файл до нижней части файла. Сразу над конечным тегом `</VersionOverrides>` вы найдете следующую разметку.

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
    > Значением **\<Resource\>** является **универсальный код** ресурса (URI) идентификатора приложения, заданный при регистрации надстройки. Этот **\<Scopes\>** раздел используется только для создания диалогового окна согласия, если надстройка продается через AppSource.

1. Сохраните и закройте файл.

### <a name="setup-for-single-tenant"></a>Настройка в однотенантном режиме

Если при регистрации надстройки вы выбрали "Учетные записи только в этом каталоге организации" для SUPPORTED **ACCOUNT TYPES** , необходимо выполнить следующие дополнительные действия по настройке.

1. Вернитесь на портал Azure и откройте колонку **Обзор** регистрации надстройки. Скопируйте **Идентификатор каталога (клиента)**.

1. В файле web.config замените "common" в значении **ida:Authority** на GUID, скопированный на предыдущем шаге.   После этого значение должно выглядеть следующим образом: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.

1. Сохраните и закройте файл web.config.

## <a name="code-the-client-side"></a>Код на стороне клиента

1. Откройте файл HomeES6.js в папке **Scripts**. В нем уже есть код.

    - Полизаполнение, которое назначает объект Office.Promise глобальному объекту window, чтобы надстройка могла работать, если в Office используется пользовательский интерфейс Internet Explorer. (Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).)
    - Назначение функции `Office.initialize` , которая, в свою очередь, назначает обработчик событию нажатия `getGraphAccessTokenButton` кнопки.
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

1. После функции `getGraphData` добавьте следующую функцию. Обратите внимание, что функция `handleClientSideErrors` будет создана позже.

    > [!NOTE]
    > Чтобы различать два маркера доступа, с помощью которые вы работаете в этой статье, маркер, возвращенный методом getAccessToken(), называется маркером начальной загрузки. Позже он обменивается через поток On-Behalf-Of для нового маркера с доступом к Microsoft Graph.

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


1. Замените `TODO 1` приведенный ниже код, чтобы получить маркер доступа от узла Office. Параметр *options* содержит следующие параметры, переданные из предыдущей функции `getGraphData()` .

    - `allowSignInPrompt` имеет значение true. В этом случае Office предложит пользователю выполнить вход, если пользователь еще не выполнил вход в Office.
    - `allowConsentPrompt` имеет значение true. В этом случае Office предложит пользователю предоставить надстройке доступ к Microsoft Azure Active Directory профиля пользователя, если согласие еще не было предоставлено. (Результирующий запрос *не* позволяет пользователю дать согласие на какие-либо области Microsoft Graph.)
    - `forMSGraphAccess` имеет значение true. В этом случае Office возвращает ошибку (код 13012), если пользователь или администратор не предоставили согласие на области Graph для надстройки. Чтобы получить доступ к Microsoft Graph, надстройка должна обменять маркер доступа на новый маркер доступа через поток "от имени". Значение `forMSGraphAccess` true помогает избежать сценария, в котором **getAccessToken()** выполняется успешно, но в дальнейшем поток "от имени" для Microsoft Graph завершается сбоем. Код на стороне клиента может реагировать на ошибку 13012, переходя на резервную систему авторизации.

    Также обратите внимание на следующий код:

    - Вы создадите функцию `getData` позже.
    - Параметр `/api/values` — это URL-адрес контроллера на стороне сервера, который будет использовать поток "от имени" для обмена маркера на новый маркер доступа для вызова Microsoft Graph.

    ```javascript
    let bootstrapToken = await Office.auth.getAccessToken(options);

    getData("/api/values", bootstrapToken);
    ```

1. После функции `getGraphData` добавьте следующее. Вот что нужно знать об этом коде:

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

1. После функции `getData` добавьте следующую функцию. Обратите внимание, что `error.code` — это число (обычно в диапазоне 13xxx).

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

1. Замените `TODO 3` приведенным ниже кодом. Во всех других случаях надстройка переходит на резервную систему авторизации. Дополнительные сведения об этих ошибках см. в разделе "Устранение неполадок единого [входа в надстройки Office"](troubleshoot-sso-in-office-add-ins.md). В этой надстройке резервная система открывает диалоговое окно, которое требует, чтобы пользователь выполнил вход, даже если пользователь уже есть.

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a>Обработка ошибок на стороне сервера

1. После функции `handleClientSideErrors` добавьте следующую функцию.

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

1. Замените `TODO 6` приведенным ниже кодом.

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

1. Замените `TODO 8` приведенным ниже кодом.

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

1. Добавьте ключевое слово `partial` в объявление класса `Startup`, если его там еще нет. Оно должно выглядеть так:

    `public partial class Startup`

1. Добавьте приведенный ниже метод в класс `Startup`. Этот метод указывает, как ПО промежуточного слоя OWIN будет проверять маркеры доступа, передаваемые ему из метода `getData` в файле Home.js на стороне клиента. Процесс вызывается при каждом вызове конечной точки веб-API, содержащей атрибут `[Authorize]`.

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. Замените `TODO 1` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Код предписывает OWIN убедиться, что аудитория, указанная в маркере начальной загрузки, которая поступает из приложения Office, должна соответствовать значению, указанному в web.config.
    - Учетные записи Майкрософт имеют идентификатор GUID издателя, который отличается от guID любого клиента организации, поэтому для поддержки обоих типов учетных записей мы не проверяем издателя.
    - Параметр `SaveSigninToken` , который `true` приводит к тому, что OWIN сохраняет необработанный маркер начальной загрузки из приложения Office. Он необходим надстройке, чтобы получить маркер доступа к Microsoft Graph в потоке "от имени".
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
    - URL-адрес, передаваемый методу, — это то, где ПО промежуточного слоя OWIN получает инструкции по получению ключа, необходимого для проверки подписи маркера начальной загрузки, полученного из приложения Office. Сегмент URL-адреса "Полномочия" предоставляется файлом web.config. Это либо строка "common", либо GUID для однотенантной надстройки.

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

1. Над строкой с объявлением `ValuesController` добавьте атрибут `[Authorize]`. Это гарантирует, что надстройка будет выполнять процесс авторизации, настроенный в последней процедуре, при каждом вызове метода контроллера. Вызывать методы контроллера можно только при наличии действительного маркера доступа к надстройке.

1. Добавьте приведенный ниже метод в `ValuesController`. Обратите внимание, что возвращаемое значение — `Task<HttpResponseMessage>`, а не `Task<IEnumerable<string>>`, которое чаще используется для метода `GET api/values`. Это побочный эффект того, что логика авторизации OAuth должна находиться в контроллере, а не в ASP.NET фильтре. Некоторые условия возникновения ошибки в этой логике требуют отправки объекта HTTP-ответа в клиент надстройки.

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

    - Ваша надстройка больше не играет роль ресурса (или аудитории), к которому приложению Office и пользователю требуется доступ. Теперь она сама является клиентом, которому необходим доступ к Microsoft Graph. `ConfidentialClientApplication` — это объект "контекста клиента" MSAL.
    - Начиная с MSAL.NET 3.x.x, `bootstrapContext` — это сам маркер начальной загрузки. 
    - Полномочия предоставляются файлом web.config. Это либо строка "common", либо GUID для однотенантной надстройки.
    - MSAL выдает `profile`ошибку, если ваш код запрашивает, что на самом деле используется только в том случае, если клиентское приложение Office получает маркер для веб-приложения надстройки. Поэтому явным образом запрашивается только `Files.Read.All`.

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
    - Значение **свойства Claims** должно быть передано клиенту, который должен передать его в приложение Office, которое затем включает его в запрос нового маркера начальной загрузки. Azure AD предложит пользователю пройти все необходимые проверки подлинности.
    - API, которые создают HTTP-ответы из исключений, не знают о свойстве **Claims**, поэтому не включают его в ответ. Нам нужно создать сообщение с ним вручную. Однако настраиваемое свойство **Message** блокирует создание свойства **ExceptionMessage**, поэтому единственный способ передать идентификатор ошибки `AADSTS50076` клиенту — добавить его в настраиваемое свойство **Message**. Код JavaScript в клиенте должен будет определить, какое свойство содержится в ответе (**Message** или **ExceptionMessage**).
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
    - Объект **MsalUiRequiredException** передается методу `SendErrorToClient`. Это гарантирует, что свойство **ExceptionMessage**, содержащее информацию об ошибке, будет включено в HTTP-отклик.

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
1. Нажмите кнопку **Получить имена файлов OneDrive**. Если вы вошли в Office с помощью учетной записи Microsoft 365 для образования или рабочей учетной записи Майкрософт, а единый вход работает должным образом, первые 10 имен файлов и папок в OneDrive для бизнеса отображаются на панели задач. Если вы не вошли в систему или в сценарии, который не поддерживает единый вход, или единый вход не работает по какой-либо причине, вам будет предложено выполнить вход. После входа в систему появятся имена файлов и папок.

### <a name="testing-the-fallback-path"></a>Тестирование резервного пути

Чтобы протестировать резервный путь авторизации, выполните следующие действия, чтобы выполнить принудительный сбой пути единого входа.

1. Добавьте следующий код в верхнюю часть метода `getDataWithToken` в HomeES6.js файле.

    ```javascript
    function MockSSOError(code) {
        this.code = code;
    }
    ```

1. Затем добавьте следующую строку в верхнюю часть `try` блока в том же методе, прямо над вызовом метода `getAccessToken`.

    ```javascript
    throw new MockSSOError("13003");
    ```

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a>Обновление надстройки при переключении в промежуточную и рабочую среду

Как и все веб-надстройки Office, когда вы будете готовы перейти на промежуточный или рабочий сервер, `localhost:44355` необходимо обновить домен в манифесте с помощью нового домена. Аналогичным образом необходимо обновить домен в web.config файла.

Так как домен отображается в регистрации AAD, `localhost:44355` необходимо обновить эту регистрацию, чтобы использовать новый домен вместо того, где он отображается.
