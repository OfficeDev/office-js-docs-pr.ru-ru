# <a name="troubleshoot-error-messages-for-single-sign-on-sso"></a>Устранение ошибок единого входа

В этой статье представлено руководство по обеспечению надежной обработки специальных условий и ошибок в надстройках Office, поддерживающих единый вход, а также устранению связанных с единым входом проблем в таких надстройках.

## <a name="debugging-tools"></a>Средства отладки

Настоятельно рекомендуем использовать во время разработки средство, которое может перехватывать и отображать HTTP-запросы от веб-службы надстройки и отклики для нее. Вот два наиболее популярных из подобных средств: 

- [Fiddler](http://www.telerik.com/fiddler), предоставляемое бесплатно ([документация](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler));
- [Charles](https://www.charlesproxy.com/), предоставляемое бесплатно в течение 30 дней ([документация](https://www.charlesproxy.com/documentation/)).

При разработке API службы вы также можете попробовать следующее:

- [Postman](http://www.getpostman.com/postman), бесплатное средство ([документация](https://www.getpostman.com/docs/)).

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>Причины и обработка ошибок в методе getAccessTokenAsync

### <a name="13000"></a>13000

Надстройка или версия Office не поддерживает API [getAccessTokenAsync](http://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync). 

- Эта версия Office не поддерживает единый вход. Необходимо использовать Office 2016 версии 1710 (сборка 8629.nnnn) или более поздняя (эту версию подписки на Office 365 иногда называют "нажми и работай"). Чтобы скачать эту версию, вам может потребоваться принять участие в программе предварительной оценки Office. Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/en-us/office-insider?tab=tab-1). 
- В манифесте надстройки отсутствует подходящий раздел [WebApplicationInfo](http://dev.office.com/reference/add-ins/manifest/webapplicationinfo).

### <a name="13001"></a>13001

Пользователь не вошел в Office. Код должен повторно вызвать метод `getAccessTokenAsync` и передать значение `forceAddAccount: true` в параметре [options](http://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync#parameters). 

### <a name="13002"></a>13002

Пользователь прервал вход или отменил согласие. 
- Если ваша надстройка предоставляет функции, не требующие входа пользователя (или предоставления согласия), то в коде следует отслеживать эту ошибку и позволять надстройке продолжать работу.
- Если надстройке необходимо, чтобы пользователь выполнил вход и дал согласие, то код должен предлагать пользователю повторить операцию, но не более одного раза. 

### <a name="13003"></a>13003

Тип пользователя не поддерживается. Пользователь не вошел в Office с помощью действительной учетной записи Майкрософт либо рабочей или учебной учетной записи. Например, это может произойти, если Office работает с учетной записью локального домена. Код должен предлагать пользователю войти в Office.

### <a name="13004"></a>13004

Недопустимый ресурс. Манифест надстройки неправильно настроен. Обновите манифест. Дополнительные сведения см. в статье [Проверка манифеста и устранение связанных с ним неполадок](troubleshoot-manifest.md).

### <a name="13005"></a>13005

Недопустимое разрешение. Как правило, это означает, что приложение Office не получило предварительные разрешения для веб-службы надстройки. Дополнительные сведения см. в разделе [Создание приложения службы](../develop/sso-in-office-add-ins.md#create-the-service-application) и статье [Регистрация надстройки в конечной точке Azure AD версии 2.0](../develop/create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (ASP.NET) или [Регистрация надстройки в конечной точке Azure AD версии 2.0](../develop/create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (Node JS). Это также может произойти, если пользователь не предоставил приложению службы разрешения на доступ к своему ресурсу `profile`.

### <a name="13006"></a>13006

Ошибка клиента. Код должен предлагать пользователю выйти и перезапустить Office.

### <a name="13007"></a>13007

Ведущему приложению Office не удалось получить маркер доступа к веб-службе надстройки.
- Убедитесь, что в регистрационных данных и манифесте надстройки указаны разрешения `openid` и `profile`. Дополнительные сведения см. в статьях [Регистрация надстройки в конечной точке Azure AD версии 2.0](../develop/create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (ASP.NET) или [Регистрация надстройки в конечной точке Azure AD версии 2.0](../develop/create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (Node JS) и [Конфигурация надстройки](../develop/create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) или [Конфигурация надстройки](../develop/create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS).
- Код должен предлагать пользователю повторить попытку позже.

### <a name="13008"></a>13008

Пользователь запустил операцию, которая вызывает метод `getAccessTokenAsync`, до завершения предыдущего вызова метода `getAccessTokenAsync`. Код должен предлагать пользователю повторить операцию после завершения предыдущей операции.

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Ошибки на стороне сервера из Azure Active Directory

### <a name="conditional-access--multifactor-authentication-errors"></a>Ошибки условного доступа и многофакторной проверки подлинности
 
При использовании некоторых конфигураций удостоверений в AAD и Office 365 некоторым ресурсам, доступным через Microsoft Graph, может потребоваться многофакторная проверка подлинности (MFA), даже если она не требуется клиенту Office 365. Когда служба AAD получает запрос на получение токена для доступа к защищенному с помощью MFA ресурсу, через поток выполнения от имени другого субъекта она возвращает веб-службе надстройки сообщение JSON, содержащее свойство `claims`. Свойство claims содержит сведения о том, какие еще факторы проверки подлинности требуются. 

Код на стороне сервера должен проверить это сообщение и ретранслировать значение свойства claims клиентскому коду. Эти сведения необходимы клиенту, так как Office обрабатывает проверку подлинности для надстроек с единым входом. Сообщение для клиента может быть кодом ошибки (например, `500 Server Error` или `401 Unauthorized`) либо находиться в тексте отклика об успешном выполнении (например, `200 OK`). В обоих случаях функция обратного вызова для клиентского AJAX-вызова веб-API надстройки должна проверять этот отклик. Если значение свойства claims было ретранслировано, то код должен повторно вызвать метод `getAccessTokenAsync` и передать значение `authChallenge: CLAIMS-STRING-HERE` в параметре [options](http://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync#parameters). Когда служба AAD обнаруживает эту строку, она предлагает пользователю указать дополнительные факторы, а затем возвращает новый маркер доступа, который будет принят в потоке выполнения от имени другого субъекта.

Ниже представлены примеры, иллюстрирующие такую обработку MFA. 

- [Единый вход с использованием ASP.NET для надстройки Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO). Библиотека MSAL, используемая в этом примере, предоставляет сообщение MFA из службы AAD в качестве исключения. Код ретранслирует его клиенту в качестве отклика `500 Server Error`. В клиентском скрипте функция обратного вызова `fail` для AJAX-вызова заново вызывает метод `getAccessTokenAsync` с параметром `authChallenge`. Обратите внимание на файлы [ValuesController.cs](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Controllers/ValuesController.cs) и [Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js).
- [Единый вход с использованием NodeJS для надстройки Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO). Сообщение MFA из службы AAD отправляется клиенту в качестве отклика об успешном выполнении. В клиентском скрипте функция обратного вызова `done` для AJAX-вызова заново вызывает метод `getAccessTokenAsync` с параметром `authChallenge`. Обратите внимание на файлы [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) и [program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).

### <a name="consent-missing-errors"></a>Ошибки, вызванные отсутствием согласия

Если в службе AAD нет записи о том, что пользователь или администратор клиента согласился предоставить надстройке доступ к ресурсу Microsoft Graph, то AAD отправит вашей веб-службе сообщение об ошибке. Код должен сообщить клиенту (например, в тексте отклика `403 Forbidden`), что нужно заново вызвать метод `getAccessTokenAsync` с параметром `forceConsent: true`.

### <a name="invalid-or-missing-scope-permission-errors"></a>Ошибки, вызванные недействительными или отсутствующими областями (разрешениями)

- Код на стороне сервера должен отправить отклик `403 Forbidden` клиенту, который должен показать пользователю понятное сообщение. Если это возможно, запишите ошибку в консоли или журнале.
- Убедитесь, что в разделе [Scopes](http://dev.office.com/reference/add-ins/manifest/scopes) манифеста надстройки указаны все необходимые разрешения. Кроме того, убедитесь, что в регистрационных данных веб-службы надстройки указаны те же разрешения. Кроме того, проверьте наличие ошибок правописания. Дополнительные сведения см. в статьях [Регистрация надстройки в конечной точке Azure AD версии 2.0](../develop/create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (ASP.NET) или [Регистрация надстройки в конечной точке Azure AD версии 2.0](../develop/create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v2-0-endpoint) (Node JS) и [Конфигурация надстройки](../develop/create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) или [Конфигурация надстройки](../develop/create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS).

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Ошибки, вызванные просроченным или недействительным токеном при вызове Microsoft Graph

Некоторые библиотеки проверки подлинности и авторизации, включая MSAL, предотвращают ошибки, связанные с просроченными токенами, по мере необходимости используя кэшированный маркер обновления. Вы также можете написать собственную систему кэширования токенов. Пример такой надстройки см. в статье [Единый вход с использованием NodeJS для надстройки Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO). Обратите внимание на файл [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts).

Но если возникает ошибка, связанная с просроченным или недействительным токеном, код должен сообщить клиенту (например, в тексте отклика `401 Unauthorized`), что нужно повторно вызвать метод `getAccessTokenAsync` и повторить вызов конечной точки веб-API надстройки. При этом будет повторен поток выполнения от имени другого субъекта, чтобы получить новый токен для Microsoft Graph. 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Ошибка, вызванная недействительным токеном при вызове Microsoft Graph

Эту ошибку необходимо обрабатывать так же, как и ошибку с просроченным токеном. См. предыдущий раздел.

### <a name="invalid-audience-error"></a>Ошибка, вызванная недействительной аудиторией

Код на стороне сервера должен отправить отклик `403 Forbidden` клиенту, а тот должен показать пользователю понятное сообщение и, возможно, также записать ошибку в консоли или журнале.

Дополнительные сведения о добавлении поддержки мультитенантности для проверки токенов см. в [примере мультитенантности в Azure](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect).
