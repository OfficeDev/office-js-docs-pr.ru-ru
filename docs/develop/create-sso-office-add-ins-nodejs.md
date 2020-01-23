---
title: Создание надстройки Office на платформе Node.js с использованием единого входа
description: Узнайте, как создать надстройку на основе Node.js, использующую единый вход Office
ms.date: 01/16/2020
localization_priority: Priority
ms.openlocfilehash: 562351011ef8aaf2ba936cceea862ebfec888b11
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265454"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a>Создание надстройки Office на платформе Node.js с использованием единого входа (предварительная версия)

Ваша веб-надстройка Office может использовать процедуру входа в Office для авторизации пользователей в надстройке и Microsoft Graph. При этом им не потребуется входить повторно. Общие сведения см. в статье [Включение единого входа в надстройке Office](sso-in-office-add-ins.md).

Из этой статьи вы узнаете, как включить единый вход в надстройке, созданной с помощью Node.js и Express. Аналогичная статья, посвященная надстройке на основе ASP.NET, — [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md).

> [!NOTE]
> В качестве альтернативы выполнения действий, описанных в этой статье, для создания надстройки Office на платформе Node.js с использованием единого входа можно использовать генератор Yeoman. Генератор Yeoman упрощает процесс создания надстройки с использованием единого входа, автоматизируя действия, необходимые для настройки единого входа в Azure, и создавая код, необходимый для его использования в надстройке. Дополнительные сведения см. в статье [Краткое руководство по использованию единого входа (SSO)](../quickstarts/sso-quickstart.md).

## <a name="prerequisites"></a>Необходимые компоненты

* [Node.js](https://nodejs.org/) (последняя версия [LTS](https://nodejs.org/about/releases))

* [Git Bash](https://git-scm.com/downloads) (или другой клиент git).

* TypeScript версии 3.6.2 или более поздней.

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* Редактор кода. Рекомендуется использовать Visual Studio Code.

* Несколько файлов и папок, сохраненных в OneDrive для бизнеса в составе подписки на Office 365.

* Подписка на Microsoft Azure. Эта надстройка требует наличия Azure Active Directory (AD). В Azure AD доступны службы идентификации, которые приложения используют для проверки подлинности и авторизации. Пробную подписку можно получить на сайте [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Настройка начального проекта

1. Клонируйте или скачайте репозиторий [Office-Add-in-NodeJS-SSO](https://github.com/officedev/office-add-in-nodejs-sso).

    > [!NOTE]
    > Существует три версии примера.  
    > * В папке **Before** находится начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы. В последующих разделах этой статьи рассматривается доработка проекта.
    > * Версия примера в папке **Complete** идентична надстройке, которую вы бы создали, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью. Чтобы использовать готовую версию, просто выполните действия, описанные в этой статье, но замените папку Before на папку Completed и пропустите разделы **Код на стороне клиента** и **Код на стороне сервера**.
    > * Версия **SSOAutoSetup** — это готовый пример, который автоматизирует большинство шагов регистрации надстройки в Azure AD и ее настройки. Используйте эту версию, если нужно быстро получить рабочую надстройку с единым входом. Просто следуйте инструкциям файла сведений в папке. На определенном этапе рекомендуется выполнить шаги ручной регистрации и настройки из этой статьи, чтобы лучше понять связь между Azure AD и надстройкой. 


1. Откройте командную строку в папке **Before**.

1. Введите в консоли команду `npm install`, чтобы установить все зависимости, указанные в файле package.json.

1. Выполните команду `npm run install-dev-certs`. При запросе нажмите **Да**, чтобы установить сертификат.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Регистрация надстройки в конечной точке Azure AD версии 2.0

1. Перейдите на страницу [регистрации приложений портала Azure](https://go.microsoft.com/fwlink/?linkid=2083908), чтобы зарегистрировать свое приложение.

1. Войдите в клиент Office 365, используя учетные данные ***администратора***. Пример: MyName@contoso.onmicrosoft.com.

1. Выберите **Новая регистрация**. На странице**Зарегистрировать приложение** задайте необходимые значения следующим образом.

    * Введите **имя** `Office-Add-in-NodeJS-SSO`.
    * Для параметра **Поддерживаемые типы учетных записей** укажите вариант **Учетные записи в любом каталоге организации и личные учетные записи Майкрософт (например, Skype, Xbox, Outlook.com)**.
    * Присвойте параметру **URI перенаправления** значение ` https://localhost:44355/dialog.html`.
    * Нажмите кнопку **Зарегистрировать**.

1. На странице **Office-Add-in-NodeJS-SSO** скопируйте и сохраните значения параметров **Идентификатор приложения (клиент)** и **Идентификатор каталога (клиент)**. Они понадобятся вам позже.

    > [!NOTE]
    > Этот идентификатор представляет собой значение аудитории, используемое, когда другие приложения, например ведущее приложение Office (PowerPoint, Word, Excel и т. д.), пытаются получить авторизованный доступ к вашему приложению. Кроме того, он используется как идентификатор клиента, когда приложение, в свою очередь, пытается получить авторизованный доступ к Microsoft Graph.

1. Выберите **Проверка подлинности** в разделе **Управление**. В разделе **Неявное представление** установите флажки **Маркер доступа** и **Токен идентификатора**. В примере используется резервная система авторизации, вызываемая при недоступности единого входа. В этой системе используется неявный поток.

1. Щелкните **Сохранить** в верхней части формы.

1. Выберите **Сертификаты и секреты** в разделе **Управление**. Нажмите кнопку **Новый секрет клиента**. Введите значение параметра **Описание**, выберите соответствующий вариант для параметра **Истекает срок действия** и нажмите кнопку **Добавить**. *Сразу скопируйте значение секрета клиента и сохраните его с идентификатором приложения* перед продолжением, так как он понадобится вам позже.

1. Выберите пункт **Предоставление API** в разделе **Управление**. Щелкните ссылку **Задать**, чтобы создать URI идентификатора приложения в формате "api://$ИД приложения GUID$", где $App ID GUID$ — **идентификатор приложения (клиента)**. Вставьте `localhost:44355/` (обратите внимание на косую черту "/" в конце) между двойной косой чертой и GUID. Весь идентификатор должен отображаться в формате `api://localhost:44355/$App ID GUID$`, например: `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`. 

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
    - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office в Интернете).

    Для каждого идентификатора сделайте следующее:

    а) Нажмите кнопку **Добавить клиентское приложение**, в открывшейся панели присвойте параметру "Идентификатор клиента" соответствующий код GUID и установите флажок `api://localhost:44355/$App ID GUID$/access_as_user`.

    б) Нажмите кнопку **Добавить приложение**.

1. Выберите пункт **Разрешения API** в разделе **Управление** и нажмите кнопку **Добавить разрешение**. В открывшейся панели выберите **Microsoft Graph** и щелкните **Делегированные разрешения**.

1. Используйте поле поиска **Выбрать разрешения**, чтобы найти нужные разрешения для надстройки. Выберите следующие параметры. Для самой надстройки требуется только первое разрешение, но разрешение `profile` необходимо, чтобы ведущее приложение Office получило маркер для веб-приложения надстройки.

    * Files.Read.All
    * profile

    > [!NOTE]
    > Разрешение `User.Read` может быть уже указано по умолчанию. Незачем запрашивать ненужные разрешения, поэтому рекомендуем снять флажок рядом с разрешением, которое не требуется вашей надстройке.

1. Установите флажок для каждого отображаемого разрешения. Выбрав нужные для надстройки разрешения, нажмите кнопку **Добавить разрешения** в нижней части панели.

1. На этой же странице нажмите кнопку **Предоставить согласие администратора для [имя клиента]** и выберите **Да** в появившемся запросе подтверждения.

## <a name="configure-the-add-in"></a>Настройка надстройки

1. Откройте папку `\Begin` в скопированном проекте в редакторе кода.

1. Откройте файл `.ENV` и используйте значения, скопированные ранее. Присвойте параметру **CLIENT_ID** значение вашего **идентификатора приложения (клиента)**, а параметру **CLIENT_SECRET** — значение секрета вашего клиента. Значения **не** должны быть заключены в кавычки. По завершении файл должен выглядеть следующим образом. 

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. Откройте файл `\public\javascripts\fallbackAuthDialog.js`. В объявлении `msalConfig` замените заполнитель $application_GUID here$ на идентификатор приложения, скопированный во время регистрации надстройки. Значение не должно быть заключено в кавычки.

1. Откройте файл манифеста надстройки manifest\manifest_local.xml и прокрутите его до конца. Над закрывающим тегом `</VersionOverrides>` вы найдете следующую часть кода:

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
    > Значение **Resource** — это **URI идентификатора приложения**, указанный при регистрации надстройки. Раздел **Scopes** используется для создания диалогового окна согласия, только если надстройка продается в AppSource.

## <a name="code-the-client-side"></a>Код на стороне клиента

### <a name="create-the-sso-logic"></a>Создание логики единого входа

1. Откройте файл `public\javascripts\ssoAuthES6.js` в редакторе кода. В нем уже есть код, обеспечивающий поддержку обещаний (даже в Internet Explorer 11), и вызов `Office.onReady` для назначения обработчика единственной кнопки надстройки.

    > [!NOTE]
    > Как следует из названия, ssoAuthES6.js использует синтаксис JavaScript ES6, так как применение `async` и `await` хорошо демонстрирует простоту API единого входа. После запуска сервера localhost этот файл будет преобразован в синтаксис ES5, чтобы пример запускался в Internet Explorer 11. 

1. Добавьте следующий код под методом Office.onReady:

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exhange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         OfficeRuntime.auth.getAccessToken call.

        }
    }
    ```

1. Замените `TODO 1` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - `OfficeRuntime.auth.getAccessToken` предписывает Office получить маркер начальной загрузки из Azure AD. Маркер начальной загрузки аналогичен маркеру идентификатора, но имеет свойство `scp` (scope) со значением `access-as-user`. Такой тип маркера веб-приложение может заменить на маркер доступа к Microsoft Graph.
    - Если параметру `allowSignInPrompt` присвоено значение true, значит при отсутствии входа пользователя Office откроет всплывающее окно входа.
    - Присвоение параметру `forMSGraphAccess` значения true сигнализирует Office, что надстройка планирует использовать маркер начальной загрузки для получения маркера доступа к Micrsoft Graph вместо его использования в качестве маркера идентификатора. Если администратор клиента не предоставил согласие на доступ надстройки к Microsoft Graph, `OfficeRuntime.auth.getAccessToken` возвращает ошибку **13012**. Надстройка может отреагировать переходом на альтернативную систему проверки подлинности. Это необходимо, так как Office может запрашивать согласие только на доступ к профилю пользователя Azure AD, а не к областям Microsoft Graph. Резервная система проверки подлинности требует повторного входа пользователя в систему, и у пользователя *может* быть запрошено согласие на доступ к областям Micrsoft Graph. Таким образом, параметр `forMSGraphAccess` обеспечивает, что надстройка не будет выполнять замену маркера, которая завершится ошибкой из-за отсутствия согласия. (Так как вы предоставили согласие администратора на предыдущем шаге, этот сценарий не возникнет для этой надстройки. Но этот параметр добавлен в любом случае, чтобы продемонстрировать рекомендацию.)

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true }); 
    ```

1. Замените `TODO 2` приведенным ниже кодом. Вы создадите метод `getGraphToken` на одном из следующих шагов.

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. Замените `TODO 3` приведенным ниже кодом. Вот что нужно знать об этом коде: 

    - Если клиент Office 365 настроен на обязательное применение многофакторной проверки подлинности, в параметр `exchangeResponse` будет включено свойство `claims` со сведениями о дополнительных обязательных факторах. В этом случае следует снова вызвать `OfficeRuntime.auth.getAccessToken` с присвоением параметру `authChallenge` значения свойства утверждений. В результате AAD предложит пользователю пройти все необходимые проверки подлинности.

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. Замените `TODO 4` приведенным ниже кодом. Вот что нужно знать об этом коде: 

    - Вы создадите метод `handleAADErrors` на одном из следующих шагов. Ошибки Azure AD возвращаются клиенту в виде откликов HTTP с кодом 200. Они не вызывают ошибки, поэтому не запускается блок `catch` метода `getGraphData`.
    - Вы создадите метод `makeGraphApiCall` на одном из следующих шагов. Он выполняет вызов AJAX к конечной точке MS Graph. Ошибки перехватываются обратным вызовом `.fail` этого вызова, а не блоком `catch` метода `getGraphData`.

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. Замените `TODO 5` приведенным ниже кодом.

    - Ошибки вызова `getAccessToken` будут иметь свойство `code` с номером ошибки (обычно в диапазоне 13xxx). Вы создадите метод `handleClientSideErrors` на одном из следующих шагов.
    - Метод `showMessage` отображает текст на панели задач.

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. Под методом `getGraphData` добавьте следующую функцию. Обратите внимание, что `/auth` — это серверный экспресс-маршрут, заменяющий маркер начальной загрузки в Azure AD на маркер доступа к Microsoft Graph.

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. Под методом `getGraphToken` добавьте следующую функцию. Обратите внимание, что `error.code` — это число (обычно в диапазоне 13xxx).

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```
1. Замените `TODO 6` приведенным ниже кодом. Дополнительные сведения об этих ошибках см. в статье [Устранение ошибок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md). 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the Web.
        showMessage("Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The OfficeRuntime.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. Замените `TODO 7` приведенным ниже кодом. Дополнительные сведения об этих ошибках см. в статье [Устранение ошибок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md). Функция `dialogFallback` вызывает альтернативную систему проверки подлинности. В этой надстройке резервная система открывает диалоговое окно, требующее входа пользователя, даже если он уже выполнил вход, и использует msal.js и неявный поток, чтобы получить маркер доступа к Microsoft Graph.

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. Добавьте указанную ниже функцию под функцией `handleClientSideErrors`. 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. Иногда срок действия маркера начальной загрузки, кэшированного в Office, не истекает в момент его проверки в Office, но истекает ко времени его попадания в Azure AD для замены. Служба Azure AD ответит ошибкой **AADSTS500133**. В этом случае надстройке следует просто рекурсивно вызвать `getGraphData`. Так как срок действия кэшированного маркера начальной загрузки истек, Office получит новый маркер из Azure AD. Поэтому замените `TODO 8` приведенным ниже кодом. 

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)       
    {
        getGraphData();
    }
    ```

1. Чтобы надстройка не вошла в бесконечный цикл вызовов `getGraphData`, она должна отслеживать число вызовов `getGraphData` и обеспечивать отсутствие повторных рекурсивных вызовов. Поэтому создайте переменную счетчика в области, которая является глобальной для функций `handleAADErrors` и `getGraphData`. Подходящее место для глобальных переменных — сразу под вызовом метода `Office.onReady`.

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. Измените структуру `if` в методе `handleAADErrors`, чтобы он:

    - увеличивал значение счетчика непосредственно перед вызовом `getGraphData`;
    - выполнял тестирование, чтобы убедиться в отсутствии повторного вызова `getGraphData`. 

    Таким образом, окончательная версия структуры `if` должна выглядеть примерно так:

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. Замените `TODO 9` приведенным ниже кодом. 

    ```javascript
    else {                
        dialogFallback();
    }
    ```

1. Сохраните и закройте файл.

### <a name="get-the-data-and-add-it-to-the-office-document"></a>Получение данных и их добавление в документ Office

1. Создайте в папке `public\javascripts` файл под названием `data.js`.

1. Добавьте указанную ниже функцию в файл. Это функция, вызываемая функцией `getGraphData` при получении маркера доступа к Microsoft Graph. 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. Замените `TODO 10` приведенным ниже кодом. Вот что нужно знать об этом коде: 

    - Этот объект является параметром метода `$.ajax`.
    - `/getuserdata` — это экспресс-маршрут на сервере надстройки, создаваемый на более позднем шаге. Он вызывает конечную точку Microsoft Graph и добавляет маркер доступа в этот вызов. 

    ```javascript
    {
        type: "GET", 
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. Замените `TODO11` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - `writeFileNamesToOfficeDocument` вставляет данные из Graph в документ Office. Он определен в файле `public\javascripts\document.js`. 
    - Если `writeFileNamesToOfficeDocument` возвращает ошибку, она начнется с сообщения "Не удалось добавить имена файлов в документ".

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () { 
        showMessage("Your data has been added to the document."); 
    })
    .catch(function (error) {        
        showMessage(error);
    });
    ```

1. Сохраните и закройте файл.

## <a name="code-the-server-side"></a>Код на стороне сервера

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a>Создание маршрутизатора проверки подлинности и логики обмена маркерами

1. Откройте файл `routes\authRoute.js` и добавьте следующую функцию маршрутизации непосредственно под операторами `require` и над оператором `module.exports`. Обратите внимание, что параметр URL-адреса `router.get` имеет значение '/'. Так как этот маршрут определен в маршрутизаторе, обрабатывающем все HTTP-запросы для URL-адреса '/auth', этот маршрут эффективно обрабатывает все запросы для '/auth'. Клиентская функция `getGraphToken`, созданная ранее, вызывает этот маршрут.  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exhange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. Замените `TODO 12` приведенным ниже кодом.

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. Замените `TODO 13` приведенным ниже кодом. Вот что нужно знать об этом коде: 

    - Это начало длинного блока `else`, но закрывающая скобка `}` не находится в конце, так как будет добавлен дополнительный код. 
    - Строка `authorization` — "носитель", за которым следует маркер начальной загрузки. Поэтому первая строка блока `else` присваивает маркер для `jwt`. (JWT означает "веб-маркер JSON".)
    - Два значения `process.env.*` — это константы, назначаемые при настройке надстройки. 
    - Параметру формы `requested_token_use` присвоено значение 'on_behalf_of'. Это указывает Azure AD, что надстройка запрашивает маркер доступа к Microsoft Graph, используя поток "от имени". Azure ответит проверкой того, что маркер начальной загрузки, назначенный параметру формы `assertion`, содержит свойство `scp` со значением `access-as-user`.
    - Параметру формы `scope` присвоено значение 'Files.Read.All', что является единственной областью Microsoft Graph, требующейся надстройке.

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. Замените `TODO 14` приведенным ниже кодом, дополняющим блок `else`. Вот что нужно знать об этом коде:

    - Константе `tenant` присвоено значение 'common', так как вы сделали надстройку мультитенатной при ее регистрации в Azure AD; в частности, когда назначили параметру **Поддерживаемые типы учетных записей** значение **Учетные записи в любом каталоге организации и персональные учетные записи Майкрософт (например, Skype, Xbox, Outlook.com)**. Если вы решили поддерживать учетные записи только в том клиенте Office 365, где зарегистрирована надстройка, в этом коде `tenant` будет указан идентификатор GUID клиента. 
    - Если при запросе POST не возникает ошибка, ответ от Azure AD преобразуется в формат JSON и отправляется клиенту. Этот объект JSON содержит свойство `access_token`, которому служба Azure AD назначила маркер доступа в Microsoft Graph.

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: form(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();
            
            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. Сохраните и закройте файл.

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a>Создание маршрута для извлечения данных из Microsoft Graph

1. Откройте файл `app.js` в корневой папке проекта. Сразу под маршрутом для '/dialog.html' добавьте следующий маршрут. Этот маршрут вызывается функцией `makeGraphApiCall`, созданной на предыдущем шаге.

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. Замените `TODO 15` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Метод `makeGraphApiCall`, вызывающий этот маршрут, добавляет маркер доступа к Microsoft Graph в HTTP-запрос в качестве заголовка с именем access_token.
    - Функция `getGraphData`определена в файле `msgraph-helper.js`. (Эта функция отличается от клиентской функции `getGraphData`, определенной в файле `ssoAuthES6.js`.)
    - Последний параметр для `queryParamsSegment` задается жестко. Если вы повторно используете этот код в рабочей надстройке и какая-либо часть `queryParamsSegment` получена из введенных пользователем данных, убедитесь, что он очищен и не может быть использован для атаки путем внедрения заголовка отклика.
    - Код сводит к минимуму данные, которые должны поступать из Microsoft Graph, указывая только нужное свойство ("name") и только первые 10 имен папок или файлов.

    ```javascript
    const graphToken = req.get('access_token');    
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. Замените `TODO 16` приведенным ниже кодом. Вот что нужно знать об этом коде:

    - Если Microsoft Graph возвращает ошибку (например, недопустимый или истекший маркер), возвращаемый объект будет содержать свойство кода со значением состояния HTTP (например, 401). Код передает ошибку клиенту. Она перехватывается обратным вызовом `.fail` метода `makeGraphApiCall`.
    - Данные Microsoft Graph включают метаданные OData и теги eTag, не требующиеся надстройке, поэтому код создает новый массив, содержащий только имена файлов для отправки клиенту.

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. Сохраните и закройте файл.

## <a name="run-the-project"></a>Запуск проекта

1. Убедитесь в наличии нескольких файлов в OneDrive, чтобы можно было проверить результаты.

1. Откройте командную строку в корне папки `\Complete`. 

1. Выполните команду `npm start`. 

1. Вам потребуется загрузить неопубликованную надстройку в приложение Office (Excel, Word или PowerPoint), чтобы протестировать ее. Инструкции зависят от вашей платформы. Ссылки на инструкции доступны в разделе [Загрузка неопубликованной надстройки Office для тестирования](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. В приложении Office на вкладке ленты **Главная** нажмите кнопку **Показать надстройку** в группе **Единый вход Node.js**, чтобы открыть надстройку области задач.

1. Нажмите кнопку **Получить имена файлов OneDrive**. Если вы выполнили вход в Office с помощью рабочей или учебной (Office 365) учетной записи либо учетной записи Майкрософт и единый вход работает надлежащим образом, первые 10 имен файлов и папок из OneDrive для бизнеса вставляются в документ. (В первый раз это может занять до 15 секунд.) Если вы не выполнили вход или используете сценарий, не поддерживающий единый вход, или единый вход не работает по какой-то причине, появится запрос на вход. После входа в систему отобразятся имена файлов и папок.

> [!NOTE]
> Если вы ранее выполняли вход в Office с использованием другого идентификатора и все еще не закрыли некоторые из открытых тогда приложений Office, Office может не сменить идентификатор (даже если кажется, что это сделано). Если это произойдет, возможен сбой при вызове Microsoft Graph или возврат данных для другого идентификатора. Чтобы избежать этого, *закройте все приложения Office*, прежде чем нажимать кнопку **Получить имена файлов OneDrive**.
