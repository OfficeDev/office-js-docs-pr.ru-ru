---
title: Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET
description: ''
ms.date: 04/15/2019
localization_priority: Priority
ms.openlocfilehash: ebcf5cd72f841f5d97093e3b5f43833e97fa9947
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914307"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a>Создание надстройки Office, в которой используется единый вход, на платформе ASP.NET (предварительная версия)

Ваша надстройка может предоставлять пользователям доступ к нескольким приложениям, используя учетные данные, введенные при входе в Office. [Как включить единый вход в надстройке Office](sso-in-office-add-ins.md)

Из этой статьи вы узнаете, как включить единый вход в надстройке, созданной с помощью ASP.NET, OWIN и MSAL для .NET.

> [!NOTE]
> Сведения о создании надстройки, в которой используется единый вход, на основе Node.js см. в [этой статье](create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Предварительные условия

* Последняя доступная версия Visual Studio 2017.

* Office 365 (версии Office, распространяемые по подписке). Последняя версия для текущего месяца и сборка из канала для участников программы предварительной оценки. Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office. Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1). Обратите внимание на то, что когда сборка будет готова для выпуска на канале Semi-annual channel, поддержка функций предварительного просмотра, включая единый вход, отключается для этой сборки.

## <a name="set-up-the-starter-project"></a>Настройка начального проекта

1. Клонируйте или скачайте репозиторий [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).

1. Перейдите в папку **Before** и откройте SLN-файл в Visual Studio. Это начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы.

    > [!NOTE]
    > В том же репозитории есть готовая версия примера. Она идентична надстройке, которую вы создадите, выполнив процедуры из этой статьи, за тем исключением, что готовый проект содержит комментарии к коду. В них нет необходимости, если вы читаете эту статью. Чтобы использовать готовую версию, просто откройте файл `sln` и выполните действия, описанные в этой статье, пропустив разделы **Код на стороне клиента** и **Код на стороне сервера**.

1. Открыв проект, выполните его сборку в Visual Studio. При этом будут установлены пакеты, указанные в файле packages.config. Это может занять от пары секунд до нескольких минут в зависимости от того, сколько пакетов хранится в локальном кэше пакетов на компьютере.

    > [!NOTE]
    > Вы увидите сообщение об ошибке, касающейся пространства имен Identity. Это побочный эффект проблемы с конфигурацией, которую вы устраните на следующем этапе. Важно то, что пакеты устанавливаются.

1. В настоящий момент версия библиотеки MSAL (Microsoft.Identity.Client), которая нужна для единого входа (версия `1.1.4-preview0002`), не включена в стандартный каталог NuGet, поэтому не указана в package.config. Ее нужно установить отдельно.

   > 1. В меню **Сервис** выберите **Диспетчер пакетов NuGet** > **Консоль диспетчера пакетов**.
   > 2. В консоли выполните указанную ниже команду. Выполнение может занять минуту или больше времени, даже при быстром подключении к Интернету. Когда все будет готово, в нижней части окна консоли отобразится такое сообщение: **"Microsoft.Identity.Client 1.1.4-preview0002" успешно установлено...**.
   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`
   > 3. В **обозревателе решений** разверните элемент **Ссылки** проекта **Office-Add-in-ASPNET-SSO-WebAPI**. Убедитесь, что в него включена библиотека **Microsoft.Identity.Client**. Если ее нет или она есть, но рядом с нею отображается значок предупреждения, удалите эту запись, а затем с помощью мастера добавления ссылок Visual Studio добавьте ссылку в сборку, указав **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**

1. Еще раз выполните сборку проекта.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Регистрация надстройки в конечной точке Azure AD версии 2.0

Следующие инструкции содержат общую информацию, поэтому их можно использовать в нескольких местах. В рамках этой статьи сделайте вот что:

- Замените заполнитель **$ADD-IN-NAME$** на `Office-Add-in-ASPNET-SSO`.
- Замените заполнитель **$FQDN-WITHOUT-PROTOCOL$** на `localhost:44355`.
- Указывая разрешения в диалоговом окне **Выбор разрешений**, установите флажки для приведенных ниже разрешений. Для самой надстройки требуется только первое разрешение, а `offline_access` и `openid` требуются для библиотеки MSAL, используемой кодом на стороне сервера. Разрешение `profile` необходимо, чтобы ведущее приложение Office получило токен для веб-приложения надстройки.
  * Files.Read.All
  * offline_access
  * openid
  * profile


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a>Предоставление надстройке разрешений администратора

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a>Конфигурация надстройки

1. В приведенной ниже строке замените заполнитель {tenant_ID} на идентификатор клиента Office 365. Если вы не скопировали идентификатор клиента при регистрации надстройки с помощью AAD, воспользуйтесь одним из способов, описанных в статье [Поиск идентификатора клиента Office 365](/onedrive/find-your-office-365-tenant-id).

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. В Visual Studio откройте файл web.config. В разделе **appSettings** есть ключи, которым необходимо назначить значения.

1. Используйте строку, составленную на шаге 1, в качестве значения ключа ida:Issuer. Убедитесь, что в значении нет пробелов.

1. Введите указанные ниже значения для соответствующих ключей.

    |Ключ|Значение|
    |:-----|:-----|
    |ida:ClientID|Идентификатор приложения, полученный во время регистрации надстройки.|
    |ida:Audience|Идентификатор приложения, полученный во время регистрации надстройки.|
    |ida:Password|Пароль, который вы получили во время регистрации надстройки.|

   Ниже показан пример того, как должны выглядеть четыре измененные вами ключа. *Обратите внимание, что параметры ClientID и Audience имеют одинаковые значения*. Вы также можете использовать один ключ для обеих целей, но вашу разметку web.config будет проще повторно использовать, если вы разделите их, так как они не всегда будут одинаковыми. Кроме того, наличие отдельных ключей позволяет считать вашу надстройку и ресурсом OAuth, связанным с ведущим приложением Office, и клиентом OAuth, связанным с Microsoft Graph.

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />

    ```

   > [!NOTE]
   > Оставьте остальные параметры в разделе **appSettings** без изменений.

1. Сохраните и закройте файл.

1. В проекте надстройки откройте файл манифеста Office-Add-in-ASPNET-SSO.xml.

1. Перейдите в конец кода файла.

1. Над закрывающим тегом `</VersionOverrides>` вы найдете следующую часть кода:

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

1. Замените заполнитель {application_GUID here} *в обоих местах* разметки идентификатором приложения, скопированным во время регистрации надстройки. Символы "{}" не входят в состав идентификатора, их не нужно вставлять. Это тот же идентификатор, который использовался для ClientID и Audience в файле web.config.

    > [!NOTE]
    > * Значение **Resource** представляет собой **универсальный код ресурса (URI) идентификатора приложения**, который вы задали, когда добавляли платформу веб-API при регистрации надстройки.
    > * Раздел **Scopes** используется для создания диалогового окна предоставления разрешений, только если надстройка продается в AppSource.

1. Откройте вкладку **Предупреждения** в **списке ошибок** в Visual Studio. Если на ней есть предупреждение о том, что `<WebApplicationInfo>` не является допустимым дочерним элементом узла `<VersionOverrides>`, это означает, что используемой вами версии Visual Studio 2017 Preview не удается распознать разметку единого входа. В качестве обходного решения в надстройке Word, Excel или PowerPoint можно выполнить указанные ниже действия. Если вы работаете с надстройкой Outlook, вы найдете решение ниже.

   - **Обходное решение для Word, Excel и Powerpoint**

        1. Закомментируйте раздел `<WebApplicationInfo>` в манифесте прямо перед завершением узла `</VersionOverrides>`.

        2. Нажмите клавишу **F5**, чтобы запустить сеанс отладки. В результате будет создана копия манифеста в следующей папке (доступ к которой проще получить в **проводнике**, чем в Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

        3. В копии манифеста удалите синтаксис комментария для раздела `<WebApplicationInfo>`.

        4. Сохраните копию манифеста.

        5. Теперь необходимо принять меры, чтобы Visual Studio не перезаписал копию манифеста, когда вы в следующий раз нажмете клавишу F5. Щелкните правой кнопкой мыши узел решения в верхней части **обозревателя решений** (но не узлы проектов).

        6. В контекстном меню выберите **Свойства**. Откроется диалоговое окно **Страницы свойств решения**.

        7. Разверните пункт **Свойства конфигурации** и щелкните **Конфигурация**.

        8. Снимите флажки **Выполнить сборку** и **Развернуть** в строке для проекта **Office-Add-in-ASPNET-SSO** (но *не* проекта **Office-Add-in-ASPNET-SSO-WebAPI**).

        9. Закройте диалоговое окно, нажав кнопку **ОК**.

   - **Обходное решение для Outlook**

        1. Найдите файл `MailAppVersionOverridesV1_1.xsd` на компьютере, используемом для разработки. Он должен находиться в том каталоге, в котором установлена среда Visual Studio, в папке `./Xml/Schemas/{lcid}`. Например, при обычной установке 32-разрядной версии VS 2017 в системе, где используется английский язык (США), полный путь будет выглядеть так: `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.

        2. Измените имя существующего файла на `MailAppVersionOverridesV1_1.old`.

        3. Скопируйте измененную версию файла в папку: [Измененная схема MailAppVersionOverrides](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. Сохраните и закройте главный файл манифеста в Visual Studio.

## <a name="code-the-client-side"></a>Код на стороне клиента

1. Откройте файл Home.js в папке **Scripts**. В нем уже есть следующий код:
    * Назначение методу `Office.initialize`, которое, в свою очередь, назначает обработчик события для нажатия кнопки `getGraphAccessTokenButton`.
    * Метод `showResult` для отображения сообщения об ошибке (или данных, возвращаемых из Microsoft Graph) в нижней части области задач.
    * Метод `logErrors` для регистрации в консоли ошибок, которые не предназначены для пользователя.

1. После назначения для метода `Office.initialize` добавьте приведенный ниже код. Вот что нужно знать об этом коде:

    * При обработке ошибок в надстройке иногда автоматически выполняется еще одна попытка получить маркер доступа с помощью другого набора параметров. Переменная счетчика `timesGetOneDriveFilesHasRun` и переменная флажка `triedWithoutForceConsent` используются, чтобы предотвратить циклическое повторение неудачных попыток получить маркер.
    * Метод `getDataWithToken` создается на следующем шаге. Обратите внимание на то, что он присваивает параметру `forceConsent` значение `false`. Дополнительные сведения см. в описании следующего шага.

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

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
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account. Other kinds of accounts, like corporate domain accounts do not work.']);
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
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```

1. Замените `TODO10` приведенным ниже кодом.

    ```javascript
    default:
        logError(result);
        break;
    ```  


1. Под методом `handleClientSideErrors` добавьте приведенный ниже метод. Этот метод обрабатывает ошибки в веб-службе надстройки при неправильном выполнении потока "от имени" или получении данных от Microsoft Graph.

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

1. Замените `TODO11` указанным ниже кодом. Обратите внимание, что для большинства ошибок `4xx`, которые веб-служба будет передавать клиентской части надстройки, в ответе будет свойство **ExceptionMessage**, содержащее номер ошибки AADSTS и другие данные. Однако, когда AAD отправляет веб-службе надстройки запрос дополнительной проверки подлинности, этот запрос содержит специальное свойство **Claims** с кодом необходимой дополнительной проверки. API ASP.NET, которые создают и отправляют HTTP-ответы клиентам, не знают об этом свойстве **Claims**, поэтому не включают его в ответ. Серверный код, который вы создадите позже, будет вручную добавлять значение **Claims** в ответ, чтобы решить эту проблему. Это значение будет находиться в свойстве **Message**, поэтому код также должен анализировать это свойство.

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. Замените `TODO12` приведенным ниже кодом. Что нужно знать об этом коде:

    * Ошибка 50076 возникает, когда Microsoft Graph требует дополнительной проверки подлинности.
    * Основное приложение Office должно получить новый маркер со значением **Claims** в качестве параметра `authChallenge`. В результате AAD предложит пользователю пройти все необходимые проверки подлинности.

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }
    ```

1. Замените `TODO13` приведенным ниже кодом. Вы замените три элемента `TODO` в этом коде с использованием *внутреннего* условного блока на следующих нескольких этапах.

    ```javascript
    else if (exceptionMessage) {

        // TODO13A: Handle the case where consent has not been granted, or has been revoked.

        // TODO13B: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO13C: Handle the case where the token that the add-in's client-side sends to it's
        //          server-side is not valid because it is missing `access_as_user` scope (permission).
    }
  
    ```


1. Замените `TODO13A` приведенным ниже кодом. (Это создает первую часть *внутреннего* условного блока). Вот что нужно знать об этом коде:

    * Ошибка 65001 означает, что доступ к Microsoft Graph не был предоставлен (или был отозван) для одного или нескольких разрешений.
    * Надстройка должна получить новый маркер с параметром `forceConsent`, имеющим значение `true`.

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
       getDataWithToken({ forceConsent: true });
    }
    ```

1. Замените `TODO13B` приведенным ниже кодом. Что нужно знать об этом коде:

    * Ошибка 70011 имеет несколько значений. Главное для этой надстройки — запрашивание недопустимого разрешения, поэтому код проверяет наличие полного описания ошибки, а не только номера.
    * Надстройка должна сообщить об ошибке.

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. Замените `TODO13C` приведенным ниже кодом. Что нужно знать об этом коде:

    * Серверный код, который вы создадите позже, отправит сообщение `Missing access_as_user`, если разрешения `access_as_user` не будет в маркере доступа, который клиент надстройки отправит в AAD для использования в потоке "от имени".
    * Надстройка должна сообщить об ошибке.

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. Замените `TODO14` приведенным ниже кодом. (Этот код является частью *внешнего* условного блока и должен следовать сразу же после закрывающихся скобок структуры, которая начинается с `else if (exceptionMessage) {` и на одном уровне отступа). Вот что нужно знать об этом коде:

    * Библиотека идентификации, которую вы будете использовать в серверном коде (MSAL), должна предотвращать отправку в Microsoft Graph устаревших и недействительных маркеров; но если это все-таки произойдет, Microsoft Graph вернет веб-службе надстройки ошибку с кодом `InvalidAuthenticationToken`. Серверный код, который вы создадите позже, передаст это сообщение клиенту надстройки.
    * В этом случае надстройка должна начать заново весь процесс проверки подлинности, сбросив счетчик и переменные флага, а затем повторно вызвать метод обработчика кнопок.

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }
    ```

1. Замените `TODO15` приведенным ниже кодом.

    ```javascript
    else {
        logError(result);
    }
    ```

1. Сохраните и закройте файл.

## <a name="code-the-server-side"></a>Код на стороне сервера

### <a name="configure-the-owin-middleware"></a>Настройка ПО промежуточного слоя OWIN

1. Откройте файл Startup.cs в корневой папке проекта.

1. Добавьте ключевое слово `partial` в объявление класса Startup, если его там еще нет. Оно должно выглядеть так:

    `public partial class Startup`

1. Добавьте приведенную ниже строку в текст метода `Configuration`. Метод `ConfigureAuth` создается позже.

    `ConfigureAuth(app);`

1. Сохраните и закройте файл.

1. Щелкните правой кнопкой мыши папку **App_Start** и выберите **Добавить > Класс**.

1. В диалоговом окне **Добавить новый элемент** введите имя файла **Startup.Auth.cs** и нажмите кнопку **Добавить**.

1. Сократите имя пространства имен в новом файле до `Office_Add_in_ASPNET_SSO_WebAPI`.

1. Убедитесь, что в начале файла есть все приведенные ниже операторы `using`.

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
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
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. Замените TODO3 приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Код сообщает OWIN о необходимости убедиться, что аудитория и поставщик маркера, указанные в маркере доступа из ведущего приложения Office (который передается путем вызова метода `getData` на стороне клиента), должны совпадать со значениями, указанными в файле web.config.
    * Если задать для свойства `SaveSigninToken` значение `true`, OWIN сохранит необработанный маркер из ведущего приложения Office. Он необходим надстройке, чтобы получить маркер доступа к Microsoft Graph в потоке "от имени".
    * ПО промежуточного слоя OWIN не проверяет разрешения. Разрешения маркера доступа, которые должны включать `access_as_user`, проверяются в контроллере.

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. Замените TODO4 приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Метод `UseOAuthBearerAuthentication` вызывается вместо более распространенного метода `UseWindowsAzureActiveDirectoryBearerAuthentication`, так как последний несовместим с конечной точкой Azure AD версии 2.
    * ПО промежуточного слоя OWIN использует URL-адрес обнаружения, передаваемый методу, чтобы получить ключ, необходимый для проверки подписи в маркере доступа, полученном из ведущего приложения Office.

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. Сохраните и закройте файл.

### <a name="create-the-apivalues-controller"></a>Создание контроллера /api/values

1. Откройте файл **Controllers\ValueController.cs**.

1. Убедитесь, что в начале файла есть приведенные ниже инструкции с `using`.

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

1. Над строкой с объявлением `ValuesController` добавьте атрибут `[Authorize]`. Это гарантирует, что надстройка будет выполнять процесс авторизации, настроенный в последней процедуре, при каждом вызове метода контроллера. Вызывать методы контроллера можно только при наличии действительного маркера доступа к надстройке.

    > [!NOTE]
    > Производственная служба веб-API на основе ASP.NET MVC должна иметь специальную логику для потока "от имени" в одном или нескольких пользовательских классах **FilterAttribute**. В этом примере логика помещается в главный контроллер, чтобы можно было легко проследить весь поток авторизации и логику получения данных. Такая же модель используется в примерах авторизации в разделе [Azure Samples](https://github.com/Azure-Samples/).

1. Добавьте приведенный ниже метод в `ValuesController`. Обратите внимание, что возвращаемое значение — `Task<HttpResponseMessage>`, а не `Task<IEnumerable<string>>`, которое чаще используется для метода `GET api/values`. Это побочный эффект нахождения пользовательской логики авторизации в контроллере: при возникновении некоторых ошибок веб-служба должна отправлять HTTP-ответ клиенту надстройки.

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

1. Замените `TODO1` приведенным ниже кодом, чтобы убедиться, что в маркере указано разрешение `access_as_user`.

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
    > Для авторизации API, который отвечает за поток выполнения от имени другого субъекта, в случае надстроек Office используйте только область `access_as_user`. Для других API в службе должны быть предусмотрены отдельные требования, касающиеся областей. Это ограничивает доступ, предоставляемый с использованием маркеров, которые получает Office.

1. Замените `TODO2` приведенным ниже кодом. Вот что нужно знать об этом коде:
    * Код преобразует необработанный маркер доступа, полученный от ведущего приложения Office, в объект `UserAssertion`, который будет передан другому методу.
    * Надстройка больше не выступает в роли ресурса (или аудитории), доступ к которому необходим ведущему приложению Office и пользователю. Теперь она сама является клиентом, которому необходим доступ к Microsoft Graph. `ConfidentialClientApplication` — это объект "контекста клиента" MSAL.
    * Третий параметр конструктора `ConfidentialClientApplication` — URL-адрес перенаправления. На самом деле он не используется в потоке "от имени", но все равно рекомендуется указывать правильный URL-адрес. С помощью четвертого и пятого параметров можно определить постоянное хранилище, которое позволяет повторно использовать действительные маркеры в разных сеансах с надстройкой. В этом примере не реализуется постоянное хранилище.
    * Для работы библиотеки MSAL требуются области `openid` и `offline_access`, но если код их избыточно запрашивает, возникает ошибка. Кроме того, ошибка возникнет, если код запросит `profile` (фактически используется только при получении ведущим приложением Office токена для веб-приложения надстройки). Поэтому явным образом запрашивается только `Files.Read.All`.

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

1. Замените `TODO3` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Для начала метод `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` проверит кэш MSAL, который находится в памяти, на наличие подходящего маркера доступа. Только в случае его отсутствия запускается поток "от имени" с конечной точкой Azure AD версии 2.
    * Если ресурс Microsoft Graph требует многофакторной проверки подлинности, а пользователь еще не предоставил соответствующие данные, AAD вызовет исключение, содержащее свойство Claims.
    * Значение свойства Claims необходимо передать клиенту, который передаст его ведущему приложению Office. Последнее добавит его в запрос на получение нового токена. AAD предложит пользователю пройти все необходимые проверки подлинности.
    * Любые исключения, отличные от типа `MsalServiceException`, не перехватываются преднамеренно, поэтому будут переданы клиенту в виде сообщений `500 Server Error`.

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

1. Замените `TODO3a` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Если ресурс Microsoft Graph требует многофакторной проверки подлинности, а пользователь еще не предоставил соответствующие данные, AAD вернет состояние "400 Bad Request" с ошибкой AADSTS50076 и свойство **Claims**. MSAL выдает исключение **MsalUiRequiredException** (наследуется от исключения **MsalServiceException**) с этой информацией. 
    * Значение свойства **Claims** необходимо передать клиенту, который передаст его ведущему приложению Office. Последнее добавит его в запрос на получение нового токена. AAD предложит пользователю пройти все необходимые проверки подлинности.
    * API, которые создают HTTP-ответы из исключений, не знают о свойстве **Claims**, поэтому не включают его в ответ. Нам нужно создать сообщение с ним вручную. Однако настраиваемое свойство **Message** блокирует создание свойства **ExceptionMessage**, поэтому единственный способ передать идентификатор ошибки `AADSTS50076` клиенту — добавить его в настраиваемое свойство **Message**. Код JavaScript в клиенте должен будет определить, какое свойство содержится в ответе (**Message** или **ExceptionMessage**).
    * Сообщение создается в формате JSON, чтобы клиентский код JavaScript мог проанализировать его с помощью известных методов объекта `JSON`.
    * Вы создадите метод `SendErrorToClient` позже. Его второй параметр — объект **Exception**. В этом случае код передает `null`, потому что включение объекта **Exception** блокирует включение свойства **Message** в создаваемый HTTP-ответ.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. Замените `TODO3b` и `TODO3c` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Если вызов AAD содержал по крайней мере одно разрешение, которое не предоставил ни пользователь, ни администратор клиента (или оно было отозвано), AAD вернет состояние "400 Bad Request" с ошибкой `AADSTS65001`. MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию. Клиент должен вызвать метод `getAccessTokenAsync` повторно, используя параметр `{ forceConsent: true }`.
    *  Если вызов AAD содержал по крайней мере одно нераспознанное разрешение, AAD вернет состояние "400 Bad Request" с ошибкой `AADSTS70011`. MSAL выдает исключение **MsalUiRequiredException**, используя эту информацию. Клиент должен сообщить об этом пользователю.
    *  Полное описание включается, так как ошибка 70011 возвращается и в других случаях, и ее следует обрабатывать в этой надстройке, только когда она означает запрос недопустимого разрешения.
    *  Объект **MsalUiRequiredException** передается методу `SendErrorToClient`. Это гарантирует, что свойство **ExceptionMessage**, содержащее информацию об ошибке, будет включено в HTTP-отклик.
    *  Сообщения нет, поэтому в качестве третьего параметра передается `null`.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. Замените `TODO3d` приведенным ниже кодом. Обратите внимание, что код повторно выдает исключение, а не передает его в собственном HTTP-ответе с состоянием **HttpStatusCode.Forbidden** (401). В результате ASP.NET отправляет собственный HTTP-ответ с состоянием "500 Server Error".

    ```csharp
    else
    {
        throw e;
    }  
    ```

1. Замените `TODO4` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Классы `GraphApiHelper` и `ODataHelper` определяются в файлах из папки **Helpers**. Класс `OneDriveItem` определяется в файле из папки **Models**. В этой статье не представлено подробное описание этих классов, так как оно не имеет отношения к авторизации и единому входу.
    * Производительность будет выше, если запрашивать у Microsoft Graph только действительно необходимые данные, поэтому в коде заданы параметры `$select` и `$top`. Первый из них показывает, что нужно только свойство name, второй — что требуются только первые три названия папок или файлов.
    * Если отправленный в Microsoft Graph токен недействителен, Microsoft Graph возвращает ошибку "401 Unauthorized" с кодом "InvalidAuthenticationToken". ASP.NET затем выдает исключение **RuntimeBinderException**. Это также происходит, когда срок действия токена истек, хотя MSAL должна предотвращать отправку таких токенов. 

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

1. Замените `TODO5` приведенным ниже кодом. Вот что нужно знать об этом коде:

    * Хотя приведенный выше код запрашивает только свойство *name* элементов OneDrive, Microsoft Graph всегда включает свойство *eTag* для элементов OneDrive. Чтобы сократить количество полезных данных, отправляемых клиенту, приведенный ниже код преобразует результаты, оставляя только имена элементов.
    * Список из трех файлов и папок OneDrive отправляется клиенту в виде HTTP-ответа "200 OK".

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

1. Добавьте приведенный ниже метод под методом Get. Вот что нужно знать об этом коде:  

    * Метод передает клиенту информацию об исключении на стороне сервера.
    * Если методу будет передано исходное исключение, конструктор HttpError включит информацию из исключения в свойство **ExceptionMessage**.  
    * Если в виде исключения будет передано значение `null`, конструктор HttpError включит параметр message в свойство **Message**. Свойства **ExceptionMessage** не будет.

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

## <a name="run-the-add-in"></a>Запуск надстройки

1. Убедитесь в наличии нескольких файлов в OneDrive, чтобы можно было проверить результаты.

1. В Visual Studio нажмите клавишу F5. Откроется PowerPoint, где на ленте **Главная** появится группа **SSO ASP.NET**.

1. Нажмите кнопку **Show Add-in** (Показать надстройку) в этой группе, чтобы увидеть пользовательский интерфейс надстройки в области задач.

1. Нажмите кнопку **Get My Files from OneDrive** (Получить мои файлы из OneDrive). Если вы не вошли в Office, вам будет предложено войти.

    > [!NOTE]
    > Если ранее вы вошли в Office, используя другой идентификатор, и не закрыли некоторые из открытых тогда приложений Office, Office может не сменить идентификатор (даже если кажется, что это сделано для PowerPoint). Если это произойдет, возможен сбой при вызове Microsoft Graph или возврат данных для другого идентификатора. Чтобы избежать этого, *закройте все приложения Office*, прежде чем нажимать кнопку **Get My Files from OneDrive** (Получить мои файлы из OneDrive).

1. После входа под кнопкой появится список файлов и папок из OneDrive. Это может занять более 15 секунд, особенно в первый раз.
