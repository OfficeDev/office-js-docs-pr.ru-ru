---
title: Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET
description: Пошаговое руководство по созданию (или преобразованию) надстройки Office с внутренней частью ASP.NET для использования единого входа (SSO).
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 5556f8486529129e5f73649722ed919899e5d87e
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641293"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET

После того как пользователи войдут в Office, ваша надстройка сможет использовать те же учетные данные для предоставления им доступа к нескольким приложениям без необходимости повторного входа. Общие сведения см. в статье [Включение единого входа в надстройке Office](sso-in-office-add-ins.md).
В этой статье описывается процесс включения единого входа в надстройку, созданной с помощью ASP.NET.

> [!NOTE]
> Сведения о создании надстройки, в которой используется единый вход, на основе Node.js см. в [этой статье](create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Предварительные требования

* Visual Studio 2019 или более поздней версии.

* [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* По крайней мере несколько файлов и папок хранятся в OneDrive для бизнеса в вашей подписке на Microsoft 365.

* Подписка на Microsoft Azure. Эта надстройка требует наличия Azure Active Directory (AD). В Azure AD доступны службы идентификации, которые приложения используют для проверки подлинности и авторизации. Пробную подписку можно получить на сайте [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Настройка начального проекта

Клонируйте или скачайте репозиторий [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).

> [!NOTE]
> Существует две версии примера.
>
> * В папке **Before** находится начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы. В последующих разделах этой статьи рассматривается доработка проекта.
> * Версия примера в папке **Complete** идентична надстройке, которую вы бы создали, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью. Чтобы использовать готовую версию, просто выполните действия, описанные в этой статье, но замените папку "Before" на папку "Complete" и пропустите разделы **Код на стороне клиента** и **Код на стороне сервера**.


## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Регистрация надстройки в конечной точке Azure AD версии 2.0

1. Перейдите на страницу [регистрации приложений портала Azure](https://go.microsoft.com/fwlink/?linkid=2083908), чтобы зарегистрировать свое приложение.

1. Выполните вход с учетными данными ***администратора*** в клиенте Microsoft 365. Пример: MyName@contoso.onmicrosoft.com.

1. Выберите **Новая регистрация**. На странице**Зарегистрировать приложение** задайте необходимые значения следующим образом.

    * Введите **имя** `Office-Add-in-ASPNET-SSO`.
    * Для параметра **Поддерживаемые типы учетных записей** укажите вариант **Учетные записи в любом каталоге организации (любой каталог Azure AD — мультитенантный) и личные учетные записи Майкрософт (например, Skype, Xbox)**. (Если вы хотите, чтобы надстройка была доступна пользователям только в клиенте, в котором вы ее регистрируете, можно выбрать вариант **Учетные записи только в этом каталоге организации…**, но вам потребуется выполнить дополнительные действия по настройке. См. раздел **Настройка в однотенантном режиме** ниже.)
    * Убедитесь, что в разделе **URI перенаправления** в раскрывающемся списке выбран пункт **Интернет**, и задайте для URI значение ` https://localhost:44355/AzureADAuth/Authorize`.
    * Нажмите кнопку **Зарегистрировать**.

1. На странице **Office-Add-in-ASPNET-SSO** скопируйте и сохраните значения для **идентификатора приложения (клиента)** и **идентификатора каталога (клиента)**. Они понадобятся вам позже.

    > [!NOTE]
    > Этот идентификатор представляет собой значение аудитории, используемое, когда другие приложения, например ведущее приложение Office (PowerPoint, Word, Excel и т. д.), пытаются получить авторизованный доступ к вашему приложению. Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.

1. В разделе **Управление** выберите **Сертификаты и секреты**. Нажмите кнопку **Новый секрет клиента**. Введите значение параметра **Описание**, выберите соответствующий вариант для параметра **Истекает срок действия** и нажмите кнопку **Добавить**. *Сразу скопируйте значение секрета клиента и сохраните его с идентификатором приложения* перед продолжением, так как он понадобится вам позже.

1. В разделе **Управление** выберите **Предоставление API**. Щелкните ссылку **Задать**, чтобы создать URI идентификатора приложения в формате "api://$ИД приложения GUID$", где $App ID GUID$ — **идентификатор приложения (клиента)**. Вставьте `localhost:44355/` (обратите внимание на знак косой черты "/", добавленный в конце) после `//` и перед GUID. Весь идентификатор должен отображаться в формате `api://localhost:44355/$App ID GUID$`, например: `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. В диалоговом окне выберите **Сохранить**.

1. Нажмите кнопку **Добавить область**. В открывшейся панели введите `access_as_user` в качестве параметра **Имя области**.

1. Для параметра **Кто может давать согласие?** установите вариант **Администраторы и пользователи**.

1. Заполните поля для настройки запросов согласия администраторов и пользователей значениями, соответствующими области `access_as_user`, позволяющей ведущему приложению Office использовать веб-интерфейсы API надстройки с такими же правами, как у текущего пользователя. Возможные варианты:

    - **Отображаемое имя согласия администратора**. Office может действовать в качестве пользователя.
    - **Описание согласия администратора**. Позволяет Office вызывать веб-API надстройки с такими же правами, как у текущего пользователя.
    - **Отображаемое имя согласия пользователя**. Office может действовать от вашего имени.
    - **Описание согласия администратора**. Позволяет Office вызывать веб-API надстройки с такими же правами, как у вас.

1. Убедитесь, что параметру **Состояние** присвоено значение **Включено**.

1. Нажмите кнопку **Добавить область**.

    > [!NOTE]
    > Доменная часть имени **области**, отображаемая непосредственно под текстовым полем, должна автоматически соответствовать URI идентификатора приложения, заданного ранее, с добавлением `/access_as_user` в конце, например: `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. В разделе **Авторизованные клиентские приложения** укажите приложения, которые необходимо авторизовать для веб-приложения надстройки. Необходимо обеспечить предварительную авторизацию для всех указанных ниже идентификаторов.

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).
    - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office).
    - `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office в Интернете).
    - `08e18876-6177-487e-b8b5-cf950c1e598c` (Office в Интернете).
    - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook в Интернете).

    Для каждого идентификатора сделайте следующее:

    а) Нажмите кнопку **Добавить клиентское приложение**, в открывшейся панели присвойте параметру "Идентификатор клиента" соответствующий код GUID и установите флажок `api://localhost:44355/$App ID GUID$/access_as_user`.

    б) Нажмите кнопку **Добавить приложение**.

1. В разделе **Управление** выберите **Разрешения API** и нажмите кнопку **Добавить разрешение**. В открывшейся панели выберите **Microsoft Graph** и щелкните **Делегированные разрешения**.

1. Используйте поле поиска **Выбрать разрешения**, чтобы найти нужные разрешения для надстройки. Выберите следующие параметры. Для самой надстройки требуется только первое разрешение, но разрешение `profile` необходимо, чтобы ведущее приложение Office получило маркер для веб-приложения надстройки. (Для надстройки требуются только разрешения Files.Read.All и profile. Остальные два необходимо запросить для библиотеки MSAL.NET.)

    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > Разрешение `User.Read` может быть уже указано по умолчанию. Незачем запрашивать ненужные разрешения, поэтому рекомендуем снять флажок рядом с разрешением, которое не требуется вашей надстройке.

1. Установите флажок для каждого отображаемого разрешения. Выбрав нужные для надстройки разрешения, нажмите кнопку **Добавить разрешения** в нижней части панели.

1. На этой же странице нажмите кнопку **Предоставить согласие администратора для [имя клиента]** и выберите **Принять** в появившемся запросе подтверждения.

    > [!NOTE]
    > После нажатия кнопки **Предоставить согласие администратора для [имя клиента]** может появиться сообщение баннера с просьбой повторить попытку через несколько минут, чтобы можно было создать запрос на продолжение. В этом случае вы можете перейти к следующему разделу, ***но не забудьте вернуться на портал и нажать эту кнопку***!

## <a name="configure-the-solution"></a>Настройка решения

1. В корне папки **Before** откройте SLN-файл решения в **Visual Studio**. В **обозревателе решений** щелкните правой кнопкой мыши верхний узел (узел решения, а не узлы проектов) и выберите **Назначить запускаемые проекты**.

1. В разделе **Общие свойства** выберите **Запускаемый проект**, а затем **Несколько запускаемых проектов**. Убедитесь, что для параметра **Действие** в обоих проектах установлено значение **Запуск** и что проект, заканчивающийся на "...WebAPI", указан в списке первым. Закройте диалоговое окно.

1. Вернувшись в **Обозреватель решений**, выберите (не используя правую кнопку мыши) проект **Office-Add-in-Microsoft-Graph-ASPNETWebAPI**. Откроется область **Свойства**. Убедитесь, что для параметра **SSL включен** задано значение **True**. Убедитесь, что **URL-адрес SSL** указан как `http://localhost:44355/`.

1. В файле web.config используйте значения, скопированные ранее. Для **ida:ClientID** и **ida:Audience** укажите **идентификатор приложения (клиента)**, для **ida:Password** — секрет клиента.

    > [!NOTE]
    > **Идентификатор приложения (клиента)** представляет собой значение аудитории, используемое, когда другие приложения, например ведущее приложение Office (PowerPoint, Word, Excel), пытаются получить авторизованный доступ к вашему приложению. Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.

1. Если вы не указали вариант "Учетные записи только в этом каталоге организации" для параметра **ПОДДЕРЖИВАЕМЫЕ ТИПЫ УЧЕТНЫХ ЗАПИСЕЙ** при регистрации настройки, сохраните и закройте файл web.config. В противном случае сохраните его, но оставьте открытым. 

1. В **обозревателе решений** выберите проект **Office-Add-in-Microsoft-Graph-ASPNET** и откройте файл манифеста надстройки Office-Add-in-ASPNET-SSO.xml, а затем прокрутите вниз до конца файла.  Над закрывающим тегом `</VersionOverrides>` вы найдете следующую разметку:

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Замените заполнитель "$application_GUID here$" *в обоих местах* разметки идентификатором приложения, скопированным при регистрации надстройки. Символы "$" не входят в состав идентификатора, их не нужно вставлять. Это тот же идентификатор, который использовался для ClientID и Audience в файле web.config.

  > [!NOTE]
  > Значение **Resource** — это **URI идентификатора приложения**, указанный при регистрации надстройки. Раздел **Scopes** используется для создания диалогового окна согласия, только если надстройка продается в AppSource.

1. Сохраните и закройте файл.

### <a name="setup-for-single-tenant"></a>Настройка в однотенантном режиме

Если вы указали вариант "Учетные записи только в этом каталоге организации" для параметра **ПОДДЕРЖИВАЕМЫЕ ТИПЫ УЧЕТНЫХ ЗАПИСЕЙ** при регистрации надстройки, необходимо выполнить дополнительные шаги настройки. 

1. Вернитесь на портал Azure и откройте колонку **Обзор** регистрации надстройки. Скопируйте **Идентификатор каталога (клиента)**.

1. В файле web.config замените "common" в значении **ida:Authority** на GUID, скопированный на предыдущем шаге.   После этого значение должно выглядеть следующим образом: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.

1. Сохраните и закройте файл web.config.

## <a name="code-the-client-side"></a>Код на стороне клиента

1. Откройте файл HomeES6.js в папке **Scripts**. В нем уже есть следующий код:

    * Полизаполнение, которое назначает объект Office.Promise глобальному объекту window, чтобы надстройка могла работать, если в Office используется пользовательский интерфейс Internet Explorer. (Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).)
    * Назначение методу `Office.initialize`, которое, в свою очередь, назначает обработчик события для нажатия кнопки `getGraphAccessTokenButton`.
    * Метод `showResult` для отображения сообщения об ошибке (или данных, возвращаемых из Microsoft Graph) в нижней части области задач.
    * Метод `logErrors` для регистрации в консоли ошибок, которые не предназначены для пользователя.
    * Код для реализации резервной системы авторизации, которая будет использоваться надстройкой в сценариях, где единый вход не поддерживается или возникла ошибка единого входа.

1. Под назначением методу `Office.initialize` добавьте приведенный ниже код. Обратите внимание на следующие особенности этого кода:

    * При обработке ошибок в надстройке иногда автоматически выполняется еще одна попытка получить маркер доступа с помощью другого набора параметров. Переменная счетчика `retryGetAccessToken` используется, чтобы предотвратить циклическое повторение неудачных попыток получить маркер.
    * Функция `getGraphData` определяется ключевым словом `async` в ES6. Синтаксис ES6 значительно упрощает использование API единого входа в надстройках Office. Это единственный файл в решении, в котором используется синтаксис, не поддерживаемый в Internet Explorer. "ES6" включается в имя файла в качестве напоминания. Компилятор TSC используется в решении для компиляции этого файла в ES5, чтобы надстройка могла работать, если в Office используется пользовательский интерфейс Internet Explorer. (См. файл tsconfig.json в корневой папке проекта.)

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, forMSGraphAccess: true });
    }
    ```

1. Добавьте указанную ниже функцию под функцией `getGraphData`. Обратите внимание, что функция `handleClientSideErrors` будет создана позже.

    ```javascript
    async function getDataWithToken() {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for an access token to Microsoft Graph and then get the data
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

1. Замените `TODO 1` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * `getAccessToken` предписывает Office получить маркер начальной загрузки из Azure AD и вернуть в надстройку.
    * `allowSignInPrompt` предписывает Office предложить пользователю выполнить вход, если он еще не вошел в Office.
    * `forMSGraphAccess` сообщает Office, что надстройка планирует заменить маркер начальной загрузки на маркер доступа к Microsoft Graph (вместо того, чтобы использовать его в качестве маркера ИД пользователя). Установка этого параметра дает Office возможность отменить процесс получения маркера начальной загрузки (и вернуть код ошибки 13012), если администратор клиента пользователя не предоставил согласие надстройке. Код на стороне клиента может реагировать на ошибку 13012, переходя на резервную систему авторизации. Если параметр не `forMSGraphAccess` используется, а администратор не предоставил согласие, то маркер начальной загрузки возвращается, но попытка обмена данными с этим продвижением приведет к ошибке. Таким образом, параметр `forMSGraphAccess` позволяет надстройке быстро перейти на резервную систему.
    * Вы создадите функцию `getData` позже.
    * Параметр `/api/values` является URL-адресом контроллера на стороне сервера, который будет осуществлять обмен маркерами и использовать маркер доступа, полученный обратно, для вызова Microsoft Graph.

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. Добавьте указанный ниже код под функцией `getGraphData`. Вот что нужно знать об этом коде:

    * Он используется и в системах единого входа, и в резервных системах авторизации.
    * Параметр `relativeUrl` является контроллером на стороне сервера.
    * Параметр `accessToken` может быть маркером начальной загрузки или маркером полного доступа.
    * `writeFileNamesToOfficeDocument` уже включен в проект.
    * Вы создадите функцию `handleServerSideErrors` позже.

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

1. Добавьте указанную ниже функцию под функцией `getData`. Обратите внимание, что `error.code` — это число (обычно в диапазоне 13xxx).

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
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
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

1. Замените `TODO 3` приведенным ниже кодом. Во всех других случаях надстройка переходит на резервную систему авторизации. Дополнительные сведения об этих ошибках можно найти [в статье Устранение неполадок единого входа в](troubleshoot-sso-in-office-add-ins.md)надстройках Office. В этой надстройке система резервного отображения открывает диалоговое окно, в котором пользователю необходимо выполнить вход, даже если он уже есть.

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a>Обработка ошибок на стороне сервера

1. Добавьте указанную ниже функцию под функцией `handleClientSideErrors`.

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
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. Замените `TODO 5` приведенным ниже кодом. Когда Microsoft Graph требует дополнительной проверки подлинности, он отправляет ошибку AADSTS50076. Она содержит сведения о дополнительном требовании в свойстве **Message.Claims**. Чтобы обработать эту ошибку, код делает вторую попытку получить маркер начальной загрузки, но в этот раз он включает запрос дополнительного фактора в виде значения параметра `authChallenge`, который предписывает Azure AD предложить пользователю пройти все требуемые проверки подлинности. 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
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

1. Замените `TODO 1` приведенным ниже кодом. Что нужно знать об этом коде:

    * Код предписывает OWIN убедиться, что аудитория, указанная в маркере начальной загрузки из ведущего приложения Office, совпадает со значением, указанным в файле web.config.
    * Учетные записи Майкрософт имеют идентификатор GUID поставщика, отличный от GUID клиента организации, поэтому для поддержки обоих типов учетных записей поставщик не проверяется.
    * Если задать для свойства `SaveSigninToken` значение `true`, OWIN сохранит необработанный маркер начальной загрузки из ведущего приложения Office. Он необходим надстройке, чтобы получить маркер доступа к Microsoft Graph в потоке "от имени".
    * ПО промежуточного слоя OWIN не проверяет области. Области маркера начальной загрузки, которые должны включать `access_as_user`, проверяются в контроллере.

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. Замените `TODO 2` приведенным ниже кодом. Что нужно знать об этом коде:

    * Метод `UseOAuthBearerAuthentication` вызывается вместо более распространенного метода `UseWindowsAzureActiveDirectoryBearerAuthentication`, так как последний несовместим с конечной точкой Azure AD версии 2.
    * ПО промежуточного слоя OWIN использует URL-адрес, передаваемый методу, чтобы получить ключ, необходимый для проверки подписи в маркере начальной загрузки, полученном из ведущего приложения Office. Сегмент URL-адреса "Полномочия" предоставляется файлом web.config. Это либо строка "common", либо GUID для однотенантной надстройки.

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

1. Добавьте приведенный ниже метод в `ValuesController`. Обратите внимание, что возвращаемое значение — `Task<HttpResponseMessage>`, а не `Task<IEnumerable<string>>`, которое чаще используется для метода `GET api/values`. Это побочный эффект того, что логика авторизации OAuth находится в контроллере, а не в фильтре ASP.NET. Некоторые условия возникновения ошибки в этой логике требуют отправки объекта HTTP-ответа в клиент надстройки.

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //        token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get the access token for Microsoft Graph.

        // TODO 4: Use the token to call Microsoft Graph.
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

    * Надстройка больше не выступает в роли ресурса (или аудитории), доступ к которому необходим ведущему приложению Office и пользователю. Теперь она сама является клиентом, которому необходим доступ к Microsoft Graph. `ConfidentialClientApplication` — это объект "контекста клиента" MSAL.
    * Начиная с MSAL.NET 3.x.x, `bootstrapContext` — это сам маркер начальной загрузки. 
    * Полномочия предоставляются файлом web.config. Это либо строка "common", либо GUID для однотенантной надстройки.
    * Для работы библиотеки MSAL требуются области `openid` и `offline_access`, но если код их избыточно запрашивает, возникает ошибка. Кроме того, ошибка возникнет, если код запросит `profile` (фактически используется только при получении ведущим приложением Office токена для веб-приложения надстройки). Поэтому явным образом запрашивается только `Files.Read.All`.

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri("https://localhost:44355")
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. Замените `TODO 3` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Для начала метод `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` проверит кэш MSAL, который находится в памяти, на наличие подходящего маркера доступа. Только в случае его отсутствия запускается поток "от имени" с конечной точкой Azure AD версии 2.
    * Любые исключения, отличные от типа `MsalServiceException`, не перехватываются преднамеренно, поэтому будут переданы клиенту в виде сообщений `500 Server Error`.

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

    * Если ресурс Microsoft Graph требует многофакторной проверки подлинности, а пользователь еще не предоставил соответствующие данные, Azure AD вернет состояние "400 Bad Request" с ошибкой `AADSTS50076` и свойство **Claims**. MSAL выдает исключение **MsalUiRequiredException** (которое наследуется от **MsalServiceException**), используя эту информацию.
    * Значение свойства **Claims** необходимо передать клиенту, который передаст его ведущему приложению Office. Последнее добавит его в запрос на получение нового маркера начальной загрузки. Azure AD предложит пользователю пройти все необходимые проверки подлинности.
    * API, которые создают HTTP-ответы из исключений, не знают о свойстве **Claims**, поэтому не включают его в ответ. Нам нужно создать сообщение с ним вручную. Однако настраиваемое свойство **Message** блокирует создание свойства **ExceptionMessage**, поэтому единственный способ передать идентификатор ошибки `AADSTS50076` клиенту — добавить его в настраиваемое свойство **Message**. Код JavaScript в клиенте должен будет определить, какое свойство содержится в ответе (**Message** или **ExceptionMessage**).
    * Сообщение создается в формате JSON, чтобы клиентский код JavaScript мог проанализировать его с помощью известных методов объекта JavaScript `JSON`.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. Замените `TODO 3b` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Если вызов Azure AD содержал по крайней мере одно разрешение, которое не предоставил ни пользователь, ни администратор клиента (или оно было отозвано), Azure AD вернет состояние "400 Bad Request" с ошибкой `AADSTS65001`. MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию.
    *  Если вызов Azure AD содержал по крайней мере одно нераспознанное разрешение, Azure AD вернет состояние "400 Bad Request" с ошибкой `AADSTS70011`. MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию.
    *  Полное описание включается, так как ошибка 70011 возвращается и в других случаях, и ее следует обрабатывать в этой надстройке, только когда она означает запрос недопустимого разрешения.
    *  Объект **MsalUiRequiredException** передается методу `SendErrorToClient`. Это гарантирует, что свойство **ExceptionMessage**, содержащее информацию об ошибке, будет включено в HTTP-отклик.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. Замените `TODO 3c` приведенным ниже кодом, чтобы обработать все остальные исключения **MsalServiceException**. Как отмечалось выше,

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

    ![Выбор ведущего приложения Office: Excel, PowerPoint или Word](../images/SelectHost.JPG)

1. Нажмите клавишу F5.
1. В приложении Office на вкладке ленты **Главная** в группе **Единый вход ASP.NET** выберите команду **Показать надстройку**, чтобы открыть надстройку области задач.
1. Нажмите кнопку **Получить имена файлов OneDrive**. Если вы выполнили вход в Office с помощью учетной записи Microsoft 365 для образовательных учреждений или рабочей учетной записи Майкрософт или учетной записи Майкрософт, и единый вход работает должным образом, первые 10 имен файлов и папок в OneDrive для бизнеса отображаются в области задач. Если вы не выполнили вход или используете сценарий, не поддерживающий единый вход, или единый вход не работает по какой-то причине, появится запрос на вход. После входа в систему отобразятся имена файлов и папок.
