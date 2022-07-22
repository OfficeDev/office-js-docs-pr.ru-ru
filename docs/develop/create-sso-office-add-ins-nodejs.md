---
title: Создание надстройки Office на платформе Node.js с использованием единого входа
description: Узнайте, как создать надстройку на Node.js, использующую единый вход Office.
ms.date: 07/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 470d6480308ed2695822aefd12e0b39b4abba32e
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958470"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>Создание надстройки Office на платформе Node.js с использованием единого входа

Ваша веб-надстройка Office может использовать процедуру входа в Office для авторизации пользователей в надстройке и Microsoft Graph. При этом им не потребуется входить повторно. Общие сведения см. в статье [Включение единого входа в надстройке Office](sso-in-office-add-ins.md).

В этой статье описывается процесс включения единого входа в надстройке. Пример создаемой надстройки состоит из двух частей. область задач, загружаемая в Microsoft Excel, и сервер среднего уровня, обрабатывающий вызовы Microsoft Graph для области задач. Сервер среднего уровня создается с Node.js и Express и предоставляет один REST API, `/getuserfilenames`который возвращает список первых 10 имен файлов в папке OneDrive пользователя. В области задач этот метод `getAccessToken()` используется для получения маркера доступа для вошедщего пользователя на сервер среднего уровня. Сервер среднего уровня использует поток On-Behalf-Of (OBO) для обмена маркера доступа на новый с доступом к Microsoft Graph. Этот шаблон можно расширить для доступа к любым данным Microsoft Graph. В области задач всегда вызывается REST API среднего уровня (передача маркера доступа), когда требуются службы Microsoft Graph. Средний уровень использует маркер, полученный через OBO, для вызова служб Microsoft Graph и возврата результатов в область задач.

Эта статья работает с надстройки, которая использует Node.js и Express. Аналогичная статья, посвященная надстройке на основе ASP.NET, — [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Необходимые компоненты

- [Node.js](https://nodejs.org/) (последняя версия [LTS](https://nodejs.org/about/releases))

- [Git Bash](https://git-scm.com/downloads) (или другой клиент git).

- Редактор кода— мы рекомендуем Visual Studio Code

- По крайней мере несколько файлов и папок, хранящихся OneDrive для бизнеса в подписке Microsoft 365

- Сборка Microsoft 365, поддерживающая [набор требований IdentityAPI 1.3](/javascript/api/requirement-sets/common/identity-api-requirement-sets). Вы можете получить бесплатную [песочницу](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) разработчика, которая предоставляет возобновляемую 90-дневную подписку Microsoft 365 E5 разработчика. Песочница разработчика включает подписку Microsoft Azure, которую можно использовать для регистрации приложений на последующих шагах в этой статье. При желании для регистрации приложений можно использовать отдельную подписку Microsoft Azure. Получите пробную подписку в [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Настройка начального проекта

1. Клонируйте или скачайте репозиторий [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO).

   > [!NOTE]
   > Существует две версии примера.
   >
   > - **Папка Begin** — это начальный проект. Пользовательский интерфейс и другие аспекты надстройки, не связанные непосредственно с единым входом и авторизацией, уже готовы. В последующих разделах этой статьи рассматривается доработка проекта.
   > - **Папка "** Завершение" содержит один и тот же пример со всеми шагами кодирования, описанными в этой статье. Чтобы использовать завершенную версию, просто следуйте инструкциям в этой статье, но замените "Begin" на "Complete" и пропустите разделы "Код на стороне клиента" и **"** Код" на стороне сервера среднего уровня.

1. Откройте командную строку в **папке Begin** .

1. Введите в консоли команду `npm install`, чтобы установить все зависимости, указанные в файле package.json.

1. Выполните команду `npm run install-dev-certs`. При запросе нажмите **Да**, чтобы установить сертификат.

## <a name="register-the-add-in-with-microsoft-identity-platform"></a>Зарегистрируйте надстройку в платформа удостоверений Майкрософт

Необходимо создать регистрацию приложения в Azure, которая представляет сервер среднего уровня. Это обеспечивает поддержку проверки подлинности, чтобы соответствующие маркеры доступа могли быть выданы коду клиента в JavaScript. Эта регистрация поддерживает как единый вход в клиенте, так и резервную проверку подлинности с помощью библиотеки проверки подлинности Майкрософт (MSAL).

1. Чтобы зарегистрировать приложение, перейдите на страницу [портал Azure - Регистрация приложений](https://go.microsoft.com/fwlink/?linkid=2083908), чтобы зарегистрировать приложение.

1. Войдите **_в клиент Microsoft_** 365 с учетными данными администратора. Пример: MyName@contoso.onmicrosoft.com.

1. Выберите **Новая регистрация**. На странице **Зарегистрировать приложение** задайте необходимые значения следующим образом.

   - Введите **имя** `Office-Add-in-NodeJS-SSO`.
   - **Задайте** для поддерживаемых типов учетных записей учетные записи в любом каталоге организации (любой каталог Azure AD — мультитенантный) и личных учетных записях **Майкрософт (например, Skype, Xbox).**
   - В разделе **URI перенаправления** задайте для платформы одностраничное приложение **(SPA)** со значением URI перенаправления `https://localhost:44355/dialog.html`.
   - Нажмите кнопку **Зарегистрировать**.

   > [!NOTE]
   > Тип приложения SPA используется только в том случае, если клиент использует MSAL для резервной проверки подлинности.

1. На странице **Office-Add-in-NodeJS-SSO** скопируйте и сохраните значения параметров **Идентификатор приложения (клиент)** и **Идентификатор каталога (клиент)**. Они понадобятся вам позже.

   > [!NOTE]
   > Этот **идентификатор приложения (клиента)** является значением аудитории, когда другие приложения, такие как клиентское приложение Office (например, PowerPoint, Word, Excel), ищут авторизованный доступ к приложению. Это также "идентификатор клиента" приложения, когда оно ищет авторизованный доступ к Microsoft Graph.

1. На крайней левой боковой панели выберите " **Проверка подлинности" в** разделе **"Управление"**. В разделе **"Неявное предоставление и гибридные** потоки" установите флажки для маркеров **доступа** и **маркеров идентификаторов**. В примере используется библиотека проверки подлинности Майкрософт (MSAL) для резервной проверки подлинности, если единый вход недоступен.

1. Выберите **Сохранить**.

1. В **разделе "** Управление" выберите **& секретов и** нажмите кнопку **"Создать секрет клиента"**. Введите значение параметра **Описание**, выберите соответствующий вариант для параметра **Истекает срок действия** и нажмите кнопку **Добавить**.

   Веб-приложение использует значение секрета **клиента** , чтобы подтвердить свою личность при запросе маркеров. _Запишите это значение для использования на следующем шаге — оно отображается только один раз._

1. На крайней левой боковой панели выберите **"Предоставить API"** в разделе **"Управление"**. Щелкните **ссылку "Задать** ". При этом будет создан URI идентификатора приложения в формате api://$App ID GUID$, где $App ИДЕНТИФИКАТОР GUID$ — это идентификатор приложения **(клиента**).

1. В созданном идентификаторе вставьте (обратите внимание на косую черту "/", `localhost:44355/` добавленную к концу) между двойной косой чертой и GUID. По завершении весь идентификатор должен иметь форму `api://localhost:44355/$App ID GUID$`, например `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`. Затем нажмите кнопку **Сохранить**.

1. Нажмите кнопку **Добавить область**. В открывшейся панели введите `access_as_user` в качестве параметра **Имя области**.

1. Для параметра **Кто может давать согласие?** установите вариант **Администраторы и пользователи**.

1. `access_as_user` Заполните поля для настройки запросов согласия администратора и пользователя значениями, подходящими для области, которая позволяет клиентским приложениям Office использовать веб-API надстройки с правами текущего пользователя. Предложения:

   - **Администратор отображаемое имя** согласия: Office может выступать в качестве пользователя.
   - **Описание согласия администратора**. Позволяет Office вызывать веб-API надстройки с такими же правами, как у текущего пользователя.
   - **Отображаемое имя согласия пользователя**: Office может действовать от имени пользователя.
   - **Описание согласия пользователя**: разрешить Office вызывать веб-API надстройки с тем же правами, что и у вас.

1. Убедитесь, что параметру **Состояние** присвоено значение **Включено**.

1. Нажмите кнопку **Добавить область**.

   > [!NOTE]
   > Доменная часть имени **области**, отображаемая непосредственно под текстовым полем, должна автоматически соответствовать URI идентификатора приложения, заданного ранее, с добавлением `/access_as_user` в конце, например: `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. В разделе **"** Авторизованные клиентские приложения" нажмите кнопку "Добавить клиентское приложение", `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`а затем на открывданной панели установите идентификатор клиента,  `api://localhost:44355/$app-id-guid$/access_as_user`а затем установите флажок "Авторизованные области".

1. Нажмите кнопку **Добавить приложение**.

   > [!NOTE]
   > Идентификатор `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` предварительно авторизует все конечные точки приложения Microsoft Office. Это также необходимо, если вы хотите поддерживать учетные записи Майкрософт (MSA) в Office для Windows и Mac. Кроме того, можно ввести соответствующее подмножество следующих идентификаторов, если по какой-либо причине вы хотите запретить авторизацию в Office на некоторых платформах. Просто оставьте идентификаторы платформ, с которых требуется отостановить авторизацию. Пользователи надстройки на этих платформах не смогут вызывать веб-API, но другие функции надстройки по-прежнему будут работать.
   >
   > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office).
   > - `93d53678-613d-4013-afc1-62e9e444a0a5` (Office в Интернете).
   > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook в Интернете).

1. На крайней левой боковой панели выберите разрешения **API** в разделе **"** Управление" и выберите **"Добавить разрешение"**. В открывшейся панели выберите **Microsoft Graph** и щелкните **Делегированные разрешения**.

1. Используйте поле поиска **Выбрать разрешения**, чтобы найти нужные разрешения для надстройки. Выберите следующие параметры. Только первое действительно требуется самой надстройке. но для `profile` получения `openid` маркера доступа с удостоверением пользователя для доступа к серверу среднего уровня приложению Office требуются разрешения и разрешения.

   - **Files.Read**
   - **profile**
   - **openid**

   > [!NOTE]
   > Разрешение `User.Read` может быть уже указано по умолчанию. Рекомендуется не запрашивать разрешения, которые не требуются, поэтому рекомендуется снять флажок для этого разрешения, если надстройка на самом деле не нужна.

1. Установите флажок для каждого отображаемого разрешения. Выбрав нужные для надстройки разрешения, нажмите кнопку **Добавить разрешения** в нижней части панели.

1. На этой же странице нажмите кнопку **Предоставить согласие администратора для [имя клиента]** и выберите **Да** в появившемся запросе подтверждения.

## <a name="configure-the-add-in"></a>Настройка надстройки

1. Откройте папку `\Begin` в скопированном проекте в редакторе кода.

1. Откройте файл `.ENV` и используйте значения, скопированные ранее при регистрации приложения **Office-Add-in-NodeJS-SSO** . Задайте значения следующим образом:

   | Имя              | Значение                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | **Идентификатор приложения (клиента) на** странице обзора регистрации приложения. |
   | **CLIENT_SECRET** | **Секрет клиента,** сохраненный на **странице & "Сертификаты** ".       |
   | **DIRECTORY_ID**  | **Идентификатор каталога (клиента) на** странице обзора регистрации приложения.   |

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

1. Замените заполнитель $app-id-guid _$" в_ обоих местах разметки идентификатором приложения, скопированным  при создании регистрации приложения **Office-Add-in-NodeJS-SSO**. Символы $не являются частью идентификатора, поэтому не включайте их. Это тот же идентификатор, который вы использовали для CLIENT_ID в . ENV-файл.

   > [!NOTE]
   > Значением **\<Resource\>** является **универсальный код** ресурса (URI) идентификатора приложения, заданный при регистрации надстройки. Этот **\<Scopes\>** раздел используется только для создания диалогового окна согласия, если надстройка продается через AppSource.

1. Откройте файл `\public\javascripts\fallback-msal\authConfig.js`. Замените заполнитель "$app-id-guid$" идентификатором приложения, сохраненным при регистрации приложения **Office-Add-in-NodeJS-SSO** , созданной ранее.

1. Сохраните изменения в файле.

## <a name="code-the-client-side"></a>Код на стороне клиента

### <a name="create-client-request-and-response-handler"></a>Создание обработчика запросов и ответов клиента

1. Откройте файл `public\javascripts\ssoAuthES6.js` в редакторе кода. В нем уже есть код, обеспечивающий поддержку обещаний (даже в Internet Explorer 11), и вызов `Office.onReady` для назначения обработчика единственной кнопки надстройки.

   > [!NOTE]
   > Как следует из названия, ssoAuthES6.js использует синтаксис JavaScript ES6, так как применение `async` и `await` хорошо демонстрирует простоту API единого входа. При запуске сервера localhost этот файл транспилируется в синтаксис ES5, чтобы пример был поддерживать Internet Explorer 11.

    Ключевой частью примера кода является запрос клиента. Клиентский запрос — это объект, который отслеживает сведения о запросе на вызов REST API на сервере среднего уровня. Это необходимо, так как состояние запроса клиента необходимо отслеживать или обновлять в следующих сценариях:

    - Единый вход повторяет попытку, когда вызов REST API завершается сбоем, так как ему требуется дополнительное согласие. Пример кода вызывается с `getAccessToken` обновленными вариантами проверки подлинности, получает необходимое согласие пользователя, а затем снова вызывает REST API. Цель состоит в том, чтобы не завершались сбоем в сценариях, где REST API требуется дополнительное согласие.
    - Единый вход завершается сбоем, и требуется резервная проверка подлинности. Маркер доступа приобретается через MSAL во всплывающем диалоговом окне. Цель заключается в том, чтобы не завершались сбоем в этом сценарии и корректно вернуться к альтернативному подходу проверки подлинности.

    Объект запроса клиента отслеживает следующие данные:

    - `authOptions` - [Параметры конфигурации проверки подлинности](/javascript/api/office/office.authoptions) для единого входа.
    - `authSSO` — true, если используется единый вход, в противном случае false.
    - `accessToken` — маркер доступа к серверу среднего уровня. Метод получения этого маркера для единого входа отличается от резервной проверки подлинности.
    - `url` — URL-адрес REST API, вызываемого на сервере среднего уровня.
    - `callbackHandler` — функция для передачи результатов вызова REST API.
    - `callbackFunction` — Функция, передаваемая клиентским запросам, когда она будет готова.

1. Чтобы инициализировать объект запроса клиента, в функции `createRequest` замените `TODO 1` приведенный ниже код.

   ```javascript
   const clientRequest = {
     authOptions: {
       allowSignInPrompt: true,
       allowConsentPrompt: true,
       forMSGraphAccess: true,
     },
     authSSO: authSSO,
     accessToken: null,
     url: url,
     callbackRESTApiHandler: restApiCallback,
     callbackFunction: callbackFunction,
   };
   ```

1. Замените `TODO 2` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Он проверяет, используется ли единый вход. Метод получения маркера доступа отличается для единого входа, чем для резервной проверки подлинности.
   - Если единый вход возвращает маркер доступа, он вызывает функцию `callbackfunction` . Для резервной проверки подлинности `dialogFallback`вызывается функция обратного вызова после входа пользователя через MSAL.

   ```javascript
   // Get access token.

   if (authSSO) {
     try {
       // Get access token from Office SSO.
       clientRequest.accessToken = await getAccessTokenFromSSO(
         clientRequest.authOptions
       );
       callbackFunction(clientRequest);
     } catch {
       // Use fallback authentication if SSO failed to get access token.
       switchToFallbackAuth(clientRequest);
     }
   } else {
     // Use fallback authentication to get access token.
     dialogFallback(clientRequest);
   }
   ```

1. В функции `getFileNameList` замените `TODO 3` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Функция вызывается `getFileNameList` , когда пользователь нажатием кнопки "Получить имена файлов **OneDrive** " в области задач.
   - Он создает клиентский запрос для отслеживания сведений о вызове, таких как URL-адрес REST API.
   - Когда REST API возвращает результат, он передается функции `handleGetFileNameResponse` . Этот обратный вызов передается в качестве параметра `createRequest` и отслеживается в `clientRequest.callbackRESTApiHandler`.
   - Код вызывает запрос `callWebServer` клиента для выполнения следующих действий и вызова REST API.

   ```javascript
   createRequest(
     "/getuserfilenames",
     handleGetFileNameResponse,
     async (clientRequest) => {
       await callWebServer(clientRequest);
     }
   );
   ```

1. В функции `handleGetFileNameResponse` замените `TODO 4` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Код передает ответ (содержащий список имен файлов) `writeFileNamesToOfficeDocument` для записи имен файлов в документ.
   - Код проверяет наличие ошибок. Если имена файлов записаны, отображается сообщение об успешном выполнении, в противном случае отображается сообщение об ошибке.

   ```javascript
   if (response != null) {
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

### <a name="get-the-sso-access-token"></a>Получение маркера доступа единого входа

1. В функции `getAccessTokenFromSSO` замените `TODO 5` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Он вызывает `Office.auth.getAccessToken` получение маркера доступа из Office.
   - При возникновении ошибки вызывается функция `handleSSOErrors` . Если ошибка не может быть обработана, вызывающий объект выдает ошибку. Это указывает вызываемой стороне на переход на резервную проверку подлинности.

   ```javascript
   try {
     // The access token returned from getAccessToken only has permissions to your middle-tier server APIs,
     // and it contains the identity claims of the signed-in user.

     const accessToken = await Office.auth.getAccessToken(authOptions);
     return accessToken;
   } catch (error) {
     let fallbackRequired = handleSSOErrors(error);
     if (fallbackRequired) throw error; // Rethrow the error and caller will switch to fallback auth.
     return null; // Returning a null token indicates no need for fallback (an explanation about the error condition was shown by handleSSOErrors).
   }
   ```

1. В функции `handleSSOErrors` замените `TODO 6` приведенным ниже кодом. Дополнительные сведения об этих ошибках см. в статье [Устранение ошибок единого входа в надстройках Office](troubleshoot-sso-in-office-add-ins.md).

   ```javascript
   let fallbackRequired = false;
   switch (err.code) {
   case 13001:
     // No one is signed into Office. If the add-in cannot be effectively used when no one
     // is logged into Office, then the first call of getAccessToken should pass the
     // `allowSignInPrompt: true` option. Since this sample does that, you should not see
     // this error.
     showMessage(
       "No one is signed into Office. But you can use many of the add-in's functions anyway. If you want to log in, press the Get OneDrive File Names button again."
     );
     break;
   case 13002:
     // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
     // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
     showMessage(
       "You can use many of the add-in's functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
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

1. Замените `TODO 7` приведенным ниже кодом. Дополнительные сведения об этих ошибках см. в разделе "Устранение неполадок единого [входа в надстройки Office"](troubleshoot-sso-in-office-add-ins.md). При любых ошибках, которые не могут быть обработаны, `true` возвращается вызываемой стороне. Это означает, что вызывающий объект должен переключиться на использование MSAL в качестве резервной проверки подлинности.

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

1. В функции `callWebServer` замените `TODO 8` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Фактический вызов AJAX будет выполнен функцией `ajaxCallToRESTApi` .
   - Эта функция попытается получить новый маркер доступа, если сервер среднего уровня вернет ошибку, указывающую на то, что истек срок действия текущего маркера.
   - Если вызов AJAX не может быть выполнен успешно, `switchToFallbackAuth` будет вызван для использования проверки подлинности MSAL вместо единого входа Office.

   ```javascript
   try {
     await ajaxCallToRESTApi(clientRequest);
   } catch (error) {
     if (error.statusText === "Internal Server Error") {
       const retryCall = handleWebServerErrors(error, clientRequest);
       if (retryCall && clientRequest.authSSO) {
         try {
           clientRequest.accessToken = await getAccessTokenFromSSO(
             clientRequest.authOptions
           );
           await ajaxCallToRESTApi(clientRequest);
         } catch {
           // If still an error go to fallback.
           switchToFallbackAuth(clientRequest);
           return;
         }
       }
     } else {
       console.log(JSON.stringify(error)); // Log any errors.
       showMessage(error.responseText);
     }
   }
   ```

1. В функции `ajaxCallToRESTApi` замените `TODO 9` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Функция явным образом повторно создает ошибки для обработки вызывающим объектом.

   ```javascript
   try {
     await $.ajax({
       type: "GET",
       url: clientRequest.url,
       headers: { Authorization: "Bearer " + clientRequest.accessToken },
       cache: false,
       success: function (data) {
         result = data;
         // Send result to the callback handler.
         clientRequest.callbackRESTApiHandler(result);
       },
     });
   } catch (error) {
     // This function explicitly requires the caller to handle any errors
     throw error;
   }
   ```

1. В функции `handleWebServerErrors` замените `TODO 10` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Ошибка возвращается сервером среднего уровня, который указывает тип ошибки и упрощает обработку здесь.
   - Для **ошибок Microsoft Graph** отобразите сообщение на панели задач.
   - Для **ошибки AADSTS500133** возвращает значение true, чтобы вызывающий объект знал, что срок действия маркера истек, и должен получить новый.
   - Для всех остальных сообщений отобразите сообщение на панели задач.

   ```javascript
   let retryCall = false;
   // Our middle-tier server returns a type to help handle the known cases.
   switch (err.responseJSON.type) {
     case "Microsoft Graph":
       // An error occurred when the middle-tier server called Microsoft Graph.
       showMessage(
         "Error from Microsoft Graph: " +
           JSON.stringify(err.responseJSON.errorDetails)
       );
       retryCall = false;
       break;
     case "Missing access_as_user":
       // The access_as_user scope was missing.
       showMessage("Error: Access token is missing the access_as_user scope.");
       retryCall = false;
       break;
     case "AADSTS500133": // expired token
       // On rare occasions the access token could expire after it was sent to the middle-tier server.
       // Microsoft identity platform will respond with
       // "The provided value for the 'assertion' is not valid. The assertion has expired."
       // Return true to indicate to caller they should refresh the token.
       retryCall = true;
       break;
     default:
       showMessage(
         "Unknown error from web server: " +
           JSON.stringify(err.responseJSON.errorDetails)
       );
       retryCall = false;
       if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
   }
   return retryCall;
   ```

Резервная проверка подлинности будет использовать библиотеку MSAL для входа пользователя. Сама надстройка является spa и использует регистрацию приложения SPA для доступа к серверу среднего уровня.

1. В функции `switchToFallbackAuth` замените `TODO 11` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Он задает глобальное значение `authSSO` false и создает новый клиентский запрос, использующий MSAL для проверки подлинности. Новый запрос имеет маркер доступа MSAL к серверу среднего уровня.
   - После создания запроса он вызывает `callWebServer` продолжение попытки успешного вызова сервера среднего уровня.

   ```javascript
   showMessage("Switching from SSO to fallback auth.");
   authSSO = false;
   // Create a new request for fallback auth.
   createRequest(
     clientRequest.url,
     clientRequest.callbackRESTApiHandler,
     async (fallbackRequest) => {
       // Hand off to call using fallback auth.
       await callWebServer(fallbackRequest);
     }
   );
   ```

## <a name="code-the-middle-tier-server"></a>Код сервера среднего уровня

Сервер среднего уровня предоставляет ИНТЕРФЕЙСЫ REST API для вызова клиента. Например, REST API `/getuserfilenames` получает список имен файлов из папки OneDrive пользователя. Для каждого вызова REST API клиенту требуется маркер доступа, чтобы убедиться, что клиент имеет доступ к своим данным. Маркер доступа обменивается на маркер Microsoft Graph через поток On-Behalf-Of (OBO). Новый маркер Microsoft Graph кэшируется библиотекой MSAL для последующих вызовов API. Он никогда не отправляется за пределы сервера среднего уровня. Дополнительные сведения см. в [разделе "Запрос маркера доступа среднего уровня"](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### <a name="create-the-route-and-implement-on-behalf-of-flow"></a>Создание маршрута и реализация потока On-Behalf-Of

1. Откройте файл и `routes\getFilesRoute.js` замените `TODO 12` его приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Он вызывает .`authHelper.validateJwt` Это гарантирует, что маркер доступа действителен и не был изменен.
   - Дополнительные сведения см. в [разделе "Проверка маркеров"](/azure/active-directory/develop/access-tokens#validating-tokens).

   ```javascript
   router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 13: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
   );
   ```

1. Замените `TODO 13` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Он запрашивает только необходимые минимальные области, например `files.read`.
   - Она использует MSAL `authHelper` для выполнения потока OBO в вызове .`acquireTokenOnBehalfOf`

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
     // TODO 14: Call Microsoft Graph to get list of filenames.
   } catch (err) {
     // TODO 15: Handle any errors.
   }
   ```

1. Замените `TODO 14` приведенным ниже кодом. Вот что нужно знать об этом коде:

   - Он создает URL-адрес для вызова Microsoft API Graph а затем выполняет вызов через функцию`getGraphData`.
   - Он возвращает ошибки, отправляя ответ HTTP 500 вместе с подробными сведениями.
   - При успешном выполнении он возвращает клиенту JSON со списком имен файлов.

   ```javascript
   // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
   // and only the top 10 folder or file names.
   const rootUrl = "/me/drive/root/children";

   // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
   // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
   // sanitized so that it cannot be used in a Response header injection attack.
   const params = "?$select=name&$top=10";

   const graphData = await getGraphData(response.accessToken, rootUrl, params);

   // If Microsoft Graph returns an error, such as invalid or expired token,
   // there will be a code property in the returned object set to a HTTP status (e.g. 401).
   // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
   if (graphData.code) {
     res.status(500).send({ type: "Microsoft Graph", errorDetails: graphData });
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

1. Замените `TODO 15` приведенным ниже кодом. Этот код проверяет, истек ли срок действия маркера, так как клиент может запросить новый маркер и вызвать его еще раз.

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(500).send({ type: "AADSTS500133", errorDetails: err });
   } else {
     res.status(500).send({ type: "Unknown", errorDetails: err });
   }
   ```

Пример должен обрабатывать как резервную проверку подлинности с помощью MSAL, так и проверку подлинности единого входа через Office. В примере сначала будет выполняться единый вход, `authSSO` а логическое значение в верхней части файла отслеживает, использует ли пример единый вход или переключил на резервную проверку подлинности.

## <a name="run-the-project"></a>Запуск проекта

1. Убедитесь в наличии нескольких файлов в OneDrive, чтобы можно было проверить результаты.

1. Откройте командную строку в корне папки `\Begin`.

1. Выполните команду, `npm install` чтобы установить все зависимости пакета.

1. Выполните команду, `npm start` чтобы запустить сервер среднего уровня.

1. Вам потребуется загрузить неопубликованную надстройку в приложение Office (Excel, Word или PowerPoint), чтобы протестировать ее. Инструкции зависят от вашей платформы. Ссылки на инструкции доступны в разделе [Загрузка неопубликованной надстройки Office для тестирования](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. В приложении Office на вкладке ленты **Главная** нажмите кнопку **Показать надстройку** в группе **Единый вход Node.js**, чтобы открыть надстройку области задач.

1. Нажмите кнопку **Получить имена файлов OneDrive**. Если вы вошли в Office с помощью учетной записи Microsoft 365 для образования или рабочей учетной записи Майкрософт, а единый вход работает должным образом, первые 10 имен файлов и папок в OneDrive для бизнеса будут вставлены в документ. (Первый раз может потребоваться до 15 секунд.) Если вы не вошли в систему или в сценарии, который не поддерживает единый вход, или единый вход не работает по какой-либо причине, вам будет предложено выполнить вход. После входа в систему появятся имена файлов и папок.

> [!NOTE]
> Если вы ранее выполняли вход в Office с использованием другого идентификатора и все еще не закрыли некоторые из открытых тогда приложений Office, Office может не сменить идентификатор (даже если кажется, что это сделано). Если это произойдет, возможен сбой при вызове Microsoft Graph или возврат данных для другого идентификатора. Чтобы избежать этого, _закройте все приложения Office_, прежде чем нажимать кнопку **Получить имена файлов OneDrive**.

## <a name="security-notes"></a>Заметки о безопасности

* В `/getuserfilenames` маршруте `getFilesroute.js` для создания вызова Microsoft Graph используется строка литерала. Если вы измените вызов таким образом, чтобы любая часть строки поступает из введенных пользователем данных, очищайте входные данные, чтобы их нельзя было использовать при атаке путем внедрения заголовка ответа.

* В `app.js` следующей политике безопасности содержимого для скриптов используется политика безопасности. Вы можете указать дополнительные ограничения в зависимости от потребностей безопасности надстройки.

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

Всегда следуйте рекомендациям по безопасности в платформа удостоверений Майкрософт [документации](/azure/active-directory/develop/).
