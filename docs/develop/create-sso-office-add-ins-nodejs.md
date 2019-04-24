---
title: Создание надстройки Office на платформе Node.js с использованием единого входа
description: ''
ms.date: 04/15/2019
localization_priority: Priority
ms.openlocfilehash: 2050f20139389ed1459cea7aba5e5e92858d00bc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448630"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a>Создание надстройки Office на платформе Node.js с использованием единого входа (предварительная версия)

Ваша веб-надстройка Office может использовать процедуру входа в Office для авторизации пользователей в надстройке и Microsoft Graph. При этом им не потребуется входить повторно. Общие сведения см. в статье [Включение единого входа в надстройке Office](sso-in-office-add-ins.md).

Из этой статьи вы узнаете, как включить единый вход в надстройке, созданной с помощью Node.js и Express.

> [!NOTE]
> Аналогичная статья, посвященная надстройке на основе ASP.NET, — [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Необходимые компоненты

* [Node и npm](https://nodejs.org/en/) версии 6.9.4 или более поздней.

* [Git Bash](https://git-scm.com/downloads) (или другой клиент git).

* TypeScript версии 2.2.2 или более поздней.

* Office 365 ( версии Office, распространяемые по подписке). Последняя версия для текущего месяца и сборка из канала для участников программы предварительной оценки. Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office. Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1). Обратите внимание на то, что когда сборка будет готова для выпуска на канале Semi-annual channel, поддержка функций предварительного просмотра, включая единый вход, отключается для этой сборки.

## <a name="set-up-the-starter-project"></a>Настройка начального проекта

1. Клонируйте или скачайте репозиторий [Office-Add-in-NodeJS-SSO](https://github.com/officedev/office-add-in-nodejs-sso).

    > [!NOTE]
    > Существует три версии примера.  
    > * В папке **Before** находится начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы. В последующих разделах этой статьи рассматривается доработка проекта.
    > * Версия примера в папке **Completed** идентична надстройке, которую вы бы создали, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью. Чтобы использовать готовую версию, просто выполните действия, описанные в этой статье, но замените папку Before на папку Completed и пропустите разделы **Код на стороне клиента** и **Код на стороне сервера**.
    > * Версия в папке **Completed Multitenant** — готовый пример, который поддерживает мультитенантность. Изучите этот пример, если вы намерены поддерживать учетные записи Майкрософт с разных доменов с единым входом.
    >
    > _Вне зависимости от используемой версии вам понадобится сделать доверенным сертификат для localhost. См. примечание "ВАЖНО!" в файле сведений о репозитории._

1. Откройте консоль Git bash в папке **Before**.

1. Введите в консоли команду `npm install`, чтобы установить все зависимости, указанные в файле package.json.

1. Введите в консоли команду `npm run build`, чтобы выполнить сборку проекта.

    > [!NOTE]
    > Могут возникать ошибки сборки с сообщениями о том, что некоторые переменные объявлены, но не используются. Игнорируйте эти ошибки. Они возникают из-за того, что в исходной версии примера отсутствует код, который будет добавлен позже.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Регистрация надстройки в конечной точке Azure AD версии 2.0

Следующие инструкции содержат общую информацию, поэтому их можно использовать в нескольких местах. В рамках этой статьи сделайте вот что:

- Замените заполнитель **$ADD-IN-NAME$** на `Office-Add-in-NodeJS-SSO`.
- Замените заполнитель **$FQDN-WITHOUT-PROTOCOL$** на `localhost:3000`.
- Указывая разрешения в диалоговом окне **Выбор разрешений**, установите флажки для приведенных ниже разрешений. Для самой надстройки требуется только первое разрешение, но разрешение `profile` необходимо, чтобы ведущее приложение Office получило маркер для веб-приложения надстройки.
  * Files.Read.All
  * profile

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a>Предоставление надстройке разрешений администратора

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a>Настройка надстройки

1. В редакторе кода откройте файл src\server.ts. В начале этого файла есть вызов конструктора класса `AuthModule`. У конструктора есть строковые параметры, которым необходимо назначить значения.

1. В свойстве `client_id` замените заполнитель `{client GUID}` на идентификатор приложения, сохраненный во время регистрации надстройки. В результате должен остаться только GUID в одиночных кавычках. Значение не должно содержать символов {}.

1. В свойстве `client_secret` замените заполнитель `{client secret}` на секрет приложения, сохраненный во время регистрации надстройки.

1. В свойстве `audience` замените заполнитель `{audience GUID}` на идентификатор приложения, сохраненный во время регистрации надстройки. (Это то же значение, которое вы назначили свойству `client_id`.)
  
1. В строке, назначенной свойству `issuer`, вы увидите заполнитель *{O365 tenant GUID}*. Замените его идентификатором клиента Office 365. Если вы не скопировали идентификатор клиента при регистрации надстройки с помощью AAD, воспользуйтесь одним из способов, описанных в статье [Поиск идентификатора клиента Office 365](/onedrive/find-your-office-365-tenant-id). В результате значение свойства `issuer` должно выглядеть примерно так:

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. Оставьте остальные параметры конструктора `AuthModule` без изменений. Сохраните и закройте файл.

1. В корневой папке проекта откройте файл манифеста Office-Add-in-NodeJS-SSO.xml.

1. Прокрутите вниз до конца файла.

1. Над последним тегом `</VersionOverrides>` вы найдете следующую разметку:

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Замените заполнитель {application_GUID here} *в обоих местах* разметки идентификатором приложения, скопированным при регистрации надстройки. (Символы "{}" не входят в состав идентификатора, поэтому их не нужно вставлять.) Это тот же идентификатор, который использовался для ClientID и Audience в файле web.config.

    > [!NOTE]
    > * Значение **Resource** представляет собой **URI идентификатора приложения**, который вы задали, когда добавляли платформу веб-API при регистрации надстройки.
    > * Раздел **Scopes** используется для создания диалогового окна согласия, только если надстройка продается в AppSource.

1. Сохраните и закройте файл.

## <a name="code-the-client-side"></a>Код на стороне клиента

1. Откройте файл program.js в папке **public**. В нем уже есть следующий код:

    * Назначение методу `Office.initialize`, которое, в свою очередь, назначает обработчик события для нажатия кнопки `getGraphAccessTokenButton`.
    * Метод `showResult` для отображения сообщения об ошибке (или данных, возвращаемых из Microsoft Graph) в нижней части области задач.
    * Метод `logErrors` для регистрации в консоли ошибок, которые не предназначены для пользователя.

1. После назначения для метода `Office.initialize` добавьте приведенный ниже код. Вот что нужно знать об этом коде:

    * При обработке ошибок в надстройке иногда автоматически выполняется еще одна попытка получить маркер доступа с помощью другого набора параметров. Переменная счетчика `timesGetOneDriveFilesHasRun`, переменные флага `triedWithoutForceConsent` и `timesMSGraphErrorReceived` используются, чтобы для пользователя не повторялись циклически неудачные попытки получить маркер.
    * Метод `getDataWithToken` создается на следующем шаге. Обратите внимание на то, что он присваивает параметру `forceConsent` значение `false`. Дополнительные сведения см. в описании следующего шага.

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }
    ```

1. Под методом `getOneDriveFiles` добавьте приведенный ниже код. Вот что нужно знать об этом коде:

    * [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) — это новый API в Office.js, позволяющий надстройке запрашивать у ведущего приложения Office (Excel, PowerPoint, Word и т. д.) маркер доступа к надстройке (для пользователя, выполнившего вход в Office). Ведущее приложение Office, в свою очередь, запрашивает маркер у конечной точки Azure AD версии 2.0. Так как вы предварительно авторизовали ведущее приложение Office для надстройки во время ее регистрации, Azure AD отправит токен.
    * Если вход в Office не выполнен, ведущее приложение Office предложит пользователю войти.
    * Параметр настроек задает для `forceConsent` значение `false`, поэтому пользователю не будет предлагаться разрешить ведущему приложению Office доступ к надстройке при каждом ее использовании. При первом запуске надстройки вызов `getAccessTokenAsync` не будет выполнен, но логика обработки ошибок, которую вы добавите на следующем этапе, автоматически выполнит повторный вызов, при этом параметру `forceConsent` будет задано значение `true`, и пользователю будет предложено согласиться. Такая процедура выполняется только в первый раз.
    * Вы создадите метод `handleClientSideErrors` позже.

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

1. Замените строку TODO1 на приведенные ниже строки. Метод `getData` и серверный маршрут /api/values создаются позже. Для конечной точки используется относительный URL-адрес, так как она должна размещаться на том же домене, что и надстройка.

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. Под методом `getOneDriveFiles` добавьте приведенный ниже код. Вот что нужно знать об этом коде:

    * Этот метод вызывает указанную конечную точку веб-API и передает ей тот же маркер доступа, который ведущее приложение Office использовало для доступа к надстройке. На стороне сервера этот маркер доступа будет использоваться в потоке "от имени" для получения маркера доступа к Microsoft Graph.
    * Вы создадите метод `handleServerSideErrors` позже.

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

### <a name="create-the-error-handling-methods"></a>Создание методов обработки ошибок

1. Под методом `getData` добавьте приведенный ниже метод. Этот метод будет обрабатывать ошибки в клиенте надстройки, когда ведущее приложение Office не сможет получить маркер доступа к веб-службе надстройки. Сообщения о таких ошибках содержат код ошибки, поэтому данный метод различает их с помощью оператора `switch`.

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

1. Замените `TODO2` приведенным ниже кодом. Ошибка 13001 возникает, если пользователь не выполнил вход или без отклика отменил запрос на предоставление 2-го фактора проверки подлинности. В обоих случаях код повторно выполняет метод `getDataWithToken` и задает параметр для принудительного запрашивания входа.

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. Замените `TODO3` приведенным ниже кодом. Ошибка 13002 возникает, когда вход или предоставление разрешений прерывается. Попросите пользователя повторить попытку, но не более одного раза.

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }
        break;
    ```

1. Замените `TODO4` приведенным ниже кодом. Ошибка 13003 возникает, когда пользователь входит под учетной записью, отличной от рабочей, учебной или личной учетной записи Майкрософт. Попросите пользователя выйти, а затем войти с помощью учетной записи поддерживаемого типа.

    ```javascript
    case 13003:
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;
    ```

    > [!NOTE]
    > Ошибка 13004 не обрабатывается при использовании этого метода, так как она должна возникать только на стадии разработки. Ее невозможно исправить с помощью кода среды выполнения, поэтому нет смысла сообщать о ней пользователю.

1. Замените `TODO5` приведенным ниже кодом. Ошибка 13005 возникает, когда Office не имеет разрешение на использование надстройки веб-службы, либо пользователь не предоставил разрешение на использование службы для `profile`.

    ```javascript
    case 13005:
        getDataWithToken({ forceConsent: true });
        break;
    ```

1. Замените `TODO6` приведенным ниже кодом. Ошибка 13006 возникает, если происходит неопределенная ошибка ведущего приложения Office, которая может свидетельствовать о его нестабильном состоянии. Попросите пользователя перезапустить Office.

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;
    ```

1. Замените `TODO7` приведенным ниже кодом. Ошибка 13007 возникает, когда нарушается взаимодействие ведущего приложения Office с AAD, из-за чего это приложение не может получить маркер доступа к веб-службе/приложению надстройки. Это может быть из-за временного сбоя сети. Попросите пользователя повторить попытку позже.

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;
    ```

1. Замените `TODO8` приведенным ниже кодом. Ошибка 13008 возникает, когда пользователь запускает операцию, которая вызывает `getAccessTokenAsync`, до завершения предыдущего вызова.

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```

1. Замените `TODO9` указанным ниже кодом. Ошибка 13009 возникает, если надстройка не поддерживает принудительное запрашивание разрешения, но выполняется вызов `getAccessTokenAsync` с установкой для параметра `forceConsent` значения `true`. Обычно в таком случае код должен автоматически повторно запустить метод `getAccessTokenAsync` с параметром, имеющим значение `false`. Но в некоторых случаях вызов метода с установкой для параметра `forceConsent` значения `true` сам по себе является автоматическим откликом на ошибку вызова метода с установкой для параметра значения `false`. В этом случае код должен не повторять попытку, а предложить пользователю выйти и войти заново.

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;


1. Replace `TODO10` with the following code.

    ```javascript
    default:
        logError(result);
        break;
    ```  

1. Под методом `handleClientSideErrors` добавьте приведенный ниже метод. Этот метод обрабатывает ошибки в веб-службе надстройки при неправильном выполнении потока "от имени" или получении данных от Microsoft Graph.

    ```javascript
    function handleServerSideErrors(result) {

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO14: Handle the case where the token that the add-in's client-side sends to its
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. Замените `TODO11` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Существуют конфигурации Azure Active Directory, согласно которым пользователю необходимо предоставить дополнительные факторы проверки подлинности для доступа к некоторым целевым объектам Microsoft Graph (например, OneDrive), даже если пользователь может войти в Office, указав всего лишь пароль. В таком случае AAD отправит отклик, содержащий номер ошибки 50076 со свойством `Claims`.
    * Ведущее приложение Office должно получить новый маркер со значением **Claims** в качестве параметра `authChallenge`. Так AAD получит команду отобразить для пользователя запрос на прохождение всех форм проверки подлинности.

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. Замените `TODO12` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем шаге*. Вот что нужно знать об этом коде:

    * Ошибка 65001 означает, что доступ к Microsoft Graph не был предоставлен (или был отозван) для одного или нескольких разрешений.
    * Надстройка должна получить новый маркер (параметру `forceConsent` должно быть задано значение `true`).

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        getDataWithToken({ forceConsent: true });
    }
    ```

1. Замените `TODO13` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем шаге*. Вот что нужно знать об этом коде:

    * Ошибка 70011 означает, что запрошена недопустимая область (разрешение). Надстройка должна сообщить об ошибке.
    * Код регистрирует любую другую ошибку с номером ошибки AAD.

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. Замените `TODO14` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем шаге*. Вот что нужно знать об этом коде:

    * Код на стороне сервера, который вы создадите на более позднем этапе, отправит сообщение, заканчивающееся на `... expected access_as_user`, если область (разрешение) `access_as_user` будет отсутствовать в маркере доступа, отправляемом клиентом надстройки в AAD для использования в потоке "от имени".
    * Надстройка должна сообщить об ошибке.

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. Замените `TODO15` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем шаге*. Вот что нужно знать об этом коде:

    * Маловероятно, чтобы в Microsoft Graph был отправлен недействительный маркер или маркер с истекшим сроком действия. Но если это произойдет, код на стороне сервера, который вы создадите на более позднем этапе, будет заканчиваться строкой `Microsoft Graph error`.
    * В этом случае надстройка должна начать заново весь процесс проверки подлинности, сбросив счетчик `timesGetOneDriveFilesHasRun` и переменные флага `timesGetOneDriveFilesHasRun`, а затем повторно вызвать метод обработчика кнопок. Но она должна сделать это только один раз. Если ситуация повторится, надстройка должна просто зарегистрировать ошибку.
    * Код зарегистрирует ошибку, если она повторится два раза подряд.

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }
    }
    ```

1. Замените `TODO16` приведенным ниже кодом *непосредственно под последней закрывающей фигурной скобкой кода, который вы добавили на предыдущем этапе*.

    ```javascript
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a>Код на стороне сервера

На стороне сервера необходимо изменить два файла.

- Файл src\auth.js предоставляет вспомогательные функции авторизации. Он уже содержит универсальные элементы, используемые в различных потоках авторизации. Нам необходимо добавить в него функции, реализующие поток "от имени".
- Файл src\server.js содержит базовые элементы, необходимые для запуска сервера и ПО промежуточного слоя express. Нам необходимо добавить в него функции, предоставляющие домашнюю страницу, и веб-API для получения данных Microsoft Graph.

### <a name="create-a-method-to-exchange-tokens"></a>Создание метода для обмена маркерами

1. Откройте файл \src\auth.ts. Добавьте приведенный ниже метод в класс `AuthModule`. Вот что нужно знать об этом коде:

    * Параметр `jwt` — это маркер доступа к приложению. В потоке "от имени" он отправляется службе AAD в обмен на маркер доступа к ресурсу.
    * Параметр scopes содержит значение по умолчанию, но в этом примере его переопределяет код вызова.
    * Указывать параметр resource не обязательно. Его не следует использовать, если [службой токенов безопасности (STS)](/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) является конечная точка AAD версии 2.0. Конечная точка версии 2.0 получает ресурс из областей и возвращает ошибку, если ресурс отправлен в HTTP-запросе.
    * Выдача исключения в блоке `catch` *не* приводит к немедленной отправке текста "500 внутренняя ошибка сервера" клиенту. Вызов кода в файле server.js захватывает данное исключение и преобразует его в сообщение об ошибке, отправляемое клиенту.

        ```typescript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource'
                                            + JSON.stringify(exception),
                                            exception);
            }
        }
        ```

1. Замените `TODO3` приведенным ниже кодом. Вот что нужно знать об этом коде:
    * Служба токенов безопасности, поддерживающая поток "от имени", ожидает определенные пары "ключ-значение" в тексте HTTP-запроса. Этот код конструирует объект, который станет текстом запроса.
    * Свойство ресурса добавляется в текст, только если методу был передан ресурс.

        ```typescript
        const v2Params = {
                client_id: this.clientId,
                client_secret: this.clientSecret,
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt,
                requested_token_use: 'on_behalf_of',
                scope: scopes.join(' ')
            };
            let finalParams = {};
            if (resource) {
                // In JavaScript we could just add the resource property to the v2Params
                // object, but that won't compile in TypeScript.
                let v1Params  = { resource: resource };  
                for(var key in v2Params) { v1Params[key] = v2Params[key]; }
                finalParams = v1Params;
            } else {
                finalParams = v2Params;
            }
        ```

1. Замените `TODO4` приведенным ниже кодом, который отправляет HTTP-запрос конечной точке маркера для службы токенов безопасности.

    ```typescript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    });
    ```

1. Замените `TODO5` приведенным ниже кодом. Обратите внимание на то, что выдача исключения *не* приводит к немедленной отправке текста "500 внутренняя ошибка сервера" клиенту. Вызов кода в файле server.js захватывает данное исключение и преобразует его в сообщение об ошибке, отправляемое клиенту.

    ```typescript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;
    }
    ```

1. Замените `TODO6` приведенным ниже кодом. Обратите внимание на то, что код сохраняет маркер доступа для ресурса и срок его действия, а не только возвращает его. В коде вызова можно обойтись без лишних вызовов службы токенов безопасности, повторно используя действительный маркер доступа к ресурсу. В следующем разделе показано, как это сделать.

    ```typescript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken;
    ```

1. Сохраните файл, но не закрывайте его.

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>Создание метода для доступа к ресурсу с помощью потока "от имени"

1. В файле src/auth.ts добавьте метод под классом `AuthModule`. Вот что нужно знать об этом коде:

    * Приведенные выше комментарии к параметрам метода `exchangeForToken` также применимы к параметрам данного метода.
    * Метод сначала проверяет постоянное хранилище на наличие действительного маркера доступа к ресурсу, срок действия которого не истечет через минуту. Он вызывает метод `exchangeForToken`, создание которого описано в предыдущем разделе, только если это необходимо.

    ```typescript
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(await resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    }
    ```

1. Сохраните и закройте файл.

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a>Создание конечных точек, предоставляющих домашнюю страницу и данные надстройки

1. Откройте файл src\server.ts.

1. Добавьте приведенный ниже метод в конец файла. Этот метод будет предоставлять домашнюю страницу надстройки. В манифесте надстройки указан URL-адрес домашней страницы.

    ```typescript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    }));
    ```

1. Добавьте приведенный ниже метод в конец файла. Этот метод будет обрабатывать все запросы к API `values`.

    ```typescript
    app.get('/api/values', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    }));
    ```

1. Замените `TODO7` приведенным ниже кодом, который проверяет маркер доступа, полученный от ведущего приложения Office. Метод `verifyJWT` определен в файле src\auth.ts. Он всегда проверяет аудиторию и издателя. С помощью необязательного параметра мы указываем на необходимость проверить, указана ли в маркере доступа область `access_as_user`. Это единственное разрешение для надстройки, необходимое пользователю и ведущему приложению Office, чтобы получить маркер доступа к Microsoft Graph с помощью потока "от имени".

    ```typescript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' });
    ```

    > [!NOTE]
    > Для авторизации API, который отвечает за поток выполнения от имени другого субъекта, в случае надстроек Office используйте только область `access_as_user`. Для других API в службе должны быть предусмотрены отдельные требования, касающиеся областей. Это ограничивает доступ, предоставляемый с использованием маркеров, которые получает Office.

1. Замените `TODO8` приведенным ниже кодом. Обратите внимание на следующие особенности этого кода:

    * Данные вызова `acquireTokenOnBehalfOf` не включают параметр ресурса, так как мы создали объект `AuthModule` (`auth`) с использованием конечной точки AAD версии 2.0, которая не поддерживает свойство ресурса.
    * Второй параметр вызова задает разрешения, необходимые надстройке, чтобы получить список файлов и папок пользователя из OneDrive. (Разрешение `profile` не запрашивается, так как оно требуется, когда ведущее приложение Office получает маркер доступа к надстройке, а не когда вы меняете этот токен на маркер доступа к Microsoft Graph.)

    ```typescript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

1. Замените `TODO9` приведенной ниже строкой. Обратите внимание на указанные ниже особенности этого кода.

    * Класс MSGraphHelper определен в файле src\msgraph-helper.ts.
    * Чтобы сократить количество возвращаемых данных, мы указываем, что нас интересуют только первые 3 элемента и свойство name.

    ```typescript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

1. Замените `TODO10` приведенным ниже кодом. Обратите внимание на то, что этот код обрабатывает ошибки "401 не санкционировано" Microsoft Graph, которые указывают на недействительный маркер или маркер с истекшим сроком действия. Вероятность такого события крайне мала, так как его должна предотвращать логика сохранения маркеров. (См. раздел **Создание метода для доступа к ресурсу с помощью потока "от имени"** выше.) Если это произойдет, код передаст клиенту ошибку с именем "Ошибка Microsoft Graph". (См. метод `handleClientSideErrors`, созданный вами в файле program.js на одном из более ранних этапов.) Код, который вы добавите в файл ODataHelper.js на одном из более поздних этапов, поможет обработать ошибки Microsoft Graph.

    ```typescript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. Замените `TODO11` приведенным ниже кодом. Обратите внимание на то, что Microsoft Graph возвращает некоторые метаданные OData и свойство **eTag** для каждого элемента, даже если запрашивается только свойство `name`. Код отправляет клиенту только имена элементов.

    ```typescript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

1. Сохраните и закройте файл.

### <a name="add-response-handling-to-the-odatahelper"></a>Добавление обработки откликов в ODataHelper

1. Откройте файл src\odata-helper.ts. Файл почти завершен. Отсутствует текст обратного вызова обработчика для события завершения запроса. Замените `TODO` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Отклик от конечной точки OData может быть сообщением об ошибке, например 401, если конечная точка запрашивает маркер доступа, а он недействителен или срок его действия истек. Но сообщение об ошибке по-прежнему является *сообщением*, а не ошибкой вызова `https.get`, поэтому строка `on('error', reject)` в конце `https.get` не запускается. Таким образом, код отличает сообщения об успешном выполнении (200) от сообщений об ошибках и отправляет объект JSON вызывающей стороне с запрошенными данными OData или информацией об ошибке.

    ```typescript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1. Замените `TODO1` приведенным ниже кодом. Обратите внимание: код предполагает, что данные возвращаются в формате JSON.

    ```typescript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1. Замените `TODO2` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Отклик с сообщением об ошибке от источника OData будет иметь аргументы statusCode и statusMessage. При этом первый из них будет присутствовать всегда, а второй — обычно. Некоторые источники OData также добавляют в текст свойство ошибки с дополнительными сведениями, например внутренними данными или конкретизирующими сообщением и кодом.
    * Объект Promise разрешен, не отклонен. `https.get` выполняется, если веб-служба вызывает конечную точку OData "сервер-сервер". Но этот вызов поступает в контексте вызова клиентом веб-API в веб-службе. Если этот "внутренний" запрос отклонен, "внешний" запрос, отправляемый клиентом веб-службе, не выполняется. Кроме того, необходимо разрешить запрос с пользовательским объектом `Error`, если стороне, вызывающей `http.get`, необходимо передать клиенту сообщения об ошибках от конечной точки OData.

    ```typescript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;

    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. Сохраните и закройте файл.

## <a name="deploy-the-add-in"></a>Развертывание надстройки

Теперь необходимо сообщить Office, где находится надстройка.

1. Создайте сетевую папку или [предоставьте общий доступ к папке через сеть](/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).

1. Поместите копию файла манифеста Office-Add-in-NodeJS-SSO.xml из корневой папки проекта в общую папку.

1. Запустите PowerPoint и откройте документ.

1. Перейдите на вкладку **Файл**, а затем выберите **Параметры**.

1. Выберите **Центр управления безопасностью**, а затем нажмите кнопку **Параметры центра управления безопасностью**.

1. Выберите пункт **Доверенные каталоги надстроек**.

1. В поле **URL-адрес каталога** введите сетевой путь к общей папке с файлом Office-Add-in-NodeJS-SSO.xml и нажмите **Добавить каталог**.

1. Установите флажок **Показать в меню** и нажмите кнопку **ОК**.

1. Появится сообщение о том, что параметры будут применены при следующем запуске Microsoft Office. Закройте PowerPoint.

## <a name="build-and-run-the-project"></a>Сборка и запуск проекта

Выполнить сборку проекта и запустить его можно двумя способами в зависимости от того, используете ли вы Visual Studio Code. В обоих случаях при каждом изменении кода автоматически выполняется повторная сборка, после чего проект запускается.

1. Если вы не используете Visual Studio Code:
   1. Откройте терминал node и перейдите к корневой папке проекта.
   1. Введите в терминале команду **npm run build**.
   1. Откройте второй терминал node и перейдите к корневой папке проекта.
   1. Введите в терминале команду **npm run start**.

1. Если используется VS Code:
   1. Откройте проект в VS Code.
   1. Нажмите клавиши CTRL+SHIFT+B, чтобы выполнить сборку проекта.
   1. Нажмите клавишу **F5**, чтобы запустить проект в сеансе отладки.


## <a name="add-the-add-in-to-an-office-document"></a>Добавление надстройки в документ Office

1. Перезапустите PowerPoint и откройте или создайте презентацию.

1. Если вкладка **Разработчик** не отображается на ленте, включите ее с помощью следующих действий:
   1. Перейдите в раздел **Файл** | **Параметры** | **Настройка ленты**.
   1. Установите флажок, чтобы включить **разработчик** в дереве имен элементов управления в правой части страницы **Настройка ленты**.
   1. Нажмите кнопку **ОК**.

1. На вкладке **Разработчик** в PowerPoint выберите **Мои надстройки**.

1. Откройте вкладку **Общая папка**.

1. Выберите **SSO NodeJS Sample** и нажмите **ОК**.

1. На ленте **Главная** появится новая группа **SSO NodeJS** с кнопкой **Show Add-in** (Показать надстройку) и значком.

## <a name="test-the-add-in"></a>Тестирование надстройки

1. Убедитесь в наличии нескольких файлов в OneDrive, чтобы можно было проверить результаты.

1. Нажмите кнопку **Show Add-in** (Показать надстройку), чтобы открыть надстройку.

1. Откроется страница приветствия. Нажмите кнопку **Get my files from OneDrive** (Получить мои файлы из OneDrive).

1. Если вы вошли в Office, под кнопкой появится список ваших файлов и папок из OneDrive. В первый раз это может занять более 15 секунд.

1. Если вы не вошли в Office, откроется всплывающее окно с предложением войти. Список файлов и папок появится через несколько секунд после входа. *Повторно нажимать кнопку не следует.*

> [!NOTE]
> Если вы ранее выполняли вход в Office с использованием другого идентификатора и все еще не закрыли некоторые из открытых тогда приложений Office, Office может не сменить идентификатор (даже если кажется, что это сделано для PowerPoint). Если это произойдет, возможен сбой при вызове Microsoft Graph или возврат данных для другого идентификатора. Чтобы избежать этого, *закройте все приложения Office*, прежде чем нажимать кнопку **Get My Files from OneDrive** (Получить мои файлы из OneDrive).
