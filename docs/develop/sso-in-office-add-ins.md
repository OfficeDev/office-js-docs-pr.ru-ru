---
title: Включение единого входа для надстроек Office
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: 1a75f7d619d2375a2f7fcb07f6afb7e0d6261ead
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579907"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Включение единого входа для надстроек Office (предварительная версия)

Пользователи выполняют вход в Office (сетевые, мобильные и компьютерные платформы) с помощью своей личной учетной записи Майкрософт, либо их учебной или рабочей учетной записи (Office 365). Также можно воспользоваться преимуществами, предлагаемыми единым входом (SSO), и использовать его для авторизации пользователя при работе с надстройкой без повторного ввода данных для входа в систему.

![Изображение, иллюстрирующее процесс входа пользователя в надстройку](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a>Статус предварительной версии

API единого входа в настоящее время поддерживается только в рамках предварительной версии. Он может использоваться разработчиками для экспериментов, но применять его в производственных надстройках не следует. Кроме того, надстройки, в которых используется единый вход, не принимаются в [AppSource](https://appsource.microsoft.com).

Не все приложения Office поддерживают предварительную версию единого входа. Она может использоваться в Word, Excel, Outlook и PowerPoint. Чтобы больше узнать о приложениях, поддерживающих в настоящее время API единого входа, см. [Набор требований IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).

### <a name="requirements-and-best-practices"></a>Требования и рекомендации

Чтобы использовать единый вход, необходимо загрузить бета-версию библиотеки JavaScript Office из `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` на стартовой HTML-странице надстройки.

При работе с надстройкой **Outlook** не забудьте включить современную проверку подлинности для клиентов Office 365. Чтобы узнать о том, как это сделать, см. [Exchange Online: как разрешить клиенту использование современной проверки подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Полагаться на единый вход, как на единственный способ проверки подлинности *не* следует. Необходимо реализовать альтернативную систему проверки подлинности, которая может применяться надстройкой при возникновении определенных ошибок. Для этой цели можно использовать пользовательские таблицы и проверку подлинности, либо воспользоваться услугми одного из поставщиков учетных данных для входа в социальные сети. Дополнительные сведения о том, как это сделать с помощью надстройки Office, можно найти в статье [Авторизация внешних служб в вашей надстройке Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). Для *Outlook* предусмотрена рекомендованная к применению система восстановления предшествующего состояния. Чтобы узнать больше, см. [Сценарий: реализация единого входа для службы в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

### <a name="how-sso-works-at-runtime"></a>Принцип работы единого входа во время выполнения

На приведенной ниже схеме показано, как работает единый вход.

![Схема, иллюстрирующая процесс единого входа](../images/sso-overview-diagram.png)

1. В надстройке JavaScript вызывает новый API Office.js [getAccessTokenAsync](#sso-api-reference). Это сообщает основному приложению Office получить маркер доступа к надстройке. В разделе [примере маркера доступа](#example-access-token).
2. Если вход в Office не выполнен, в ведущем приложении открывается всплывающее окно, в котором пользователю предлагается войти.
3. Если пользователь запускает надстройку в первый раз, ему предлагается дать согласие.
4. Ведущее приложение Office запрашивает **маркер надстройки** у конечной точки Azure AD версии 2.0 для текущего пользователя.
5. Azure AD отправляет маркер надстройки ведущему приложению Office.
6. Ведущее приложение Office отправляет **маркер** надстройке в составе объекта результата, возвращенного при вызове метода `getAccessTokenAsync`.
7. JavaScript в надстройке может проанализировать маркер и извлечь из него требуемую информацию, например, адрес электронной почты пользователя. 
8. В качестве альтернативного варианта, надстройка может отправить HTTP-запрос своему серверу для получения дополнительных данных о пользователе, например, сведений о пользовательских настройках. Кроме того, серверу может передаваться сам маркер доступа для последующего анализа и проверки. 

## <a name="develop-an-sso-add-in"></a>Разработка надстройки с единым входом

В этом разделе описываются задачи, решаемые при создании надстройки Office с единым входом. При описании этих задач применяемые язык и платформа не учитываются. Чтобы ознакомиться с примерами подробных пошаговых руководств см.:

* [Создание надстройки Office на платформе Node.js с использованием единого входа](create-sso-office-add-ins-nodejs.md)
* [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>Создание приложения-службы

Зарегистрируйте надстройку на портале регистрации для конечной точки Azure версии 2.0: https://apps.dev.microsoft.com. Этот процесс занимает 5–10 минут и включает следующие задачи.

* Получение идентификатора клиента и секрета клиента для надстройки.
* Укажите разрешения, необходимые вашей надстройке, в конечной точке AAD версии 2.0 (и, при необходимости, в Microsoft Graph). Разрешение «Профиль» требуется всегда.
* Установка доверия надстройке для ведущего приложения Office.
* предварительная авторизация ведущего приложения Office для надстройки с помощью заданного по умолчанию разрешения *access_as_user*.

Для получения дополнительной информации об этом процессе см. статью [Регистрация надстройки Office, использующей единый вход с конечной точкой Azure AD версии 2.0](register-sso-add-in-aad-v2.md).

### <a name="configure-the-add-in"></a>Конфигурация надстройки

Добавьте новую разметку в манифест надстройки:

* **WebApplicationInfo** — родительский элемент для указанных ниже элементов.
* **Id** — идентификатор клиента надстройки. Это идентификатор приложения, предоставляемый в процессе регистрации надстройки. См. [Регистрация использующей единый вход надстройки Office в конечной точке Azure AD версии 2.0](register-sso-add-in-aad-v2.md).
* **Resource** — URL-адрес надстройки.
* **Scopes** — родительский элемент одного или нескольких элементов **Scope**.
* **Scope** — характеризует разрешение, необходимое надстройке в AAD. `profile` Это разрешение требуется всегда, а в случае, если надстройка не имеет доступа к Microsoft Graph, могут потребоваться и другие разрешения. При наличии у надстройки доступа к Microsoft Graph элементы **Scope** также требуются для необходимых разрешений Microsoft Graph, например, для `User.Read` и `Mail.Read`. Библиотекам, применяемым в коде для доступа к Microsoft Graph, могут потребоваться дополнительные разрешения. Так, к примеру, библиотеке проверки подлинности Майкрософт (MSAL) для .NET требуется разрешение `offline_access`. Чтобы больше узнать, см. [Авторизация в Microsoft Graph из надстройки Office](authorize-to-microsoft-graph.md).

Для отличных от Outlook ведущих приложений Office добавляйте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Для Outlook добавляйте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

Ниже приведен пример части кода.

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### <a name="add-client-side-code"></a>Добавление кода для клиента

Добавьте в надстройку код JavaScript для:

* Вызов [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).

* анализа маркера доступа или его передачи в серверный код надстройки; 

Далее представлен простой пример вызова `getAccessTokenAsync`. 

> [!NOTE]
> В этом примере явным образом обрабатывается только один тип ошибки. Примеры более сложной обработки ошибок см. в [Home.js в Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) и [program.js в Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). См. также [Устранение неполадок при появлении сообщений об ошибках для единого входа (SSO)](troubleshoot-sso-in-office-add-ins.md).
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

Ниже приводится простой пример передачи маркера надстройки серверу. Данный маркер используется в качестве заголовка `Authorization` при отправке запроса серверу. В этом примере демонстрируется отправка данных JSON: здесь используется метод `POST`, но для отправки маркера доступа без выполнения записи на сервер достаточно применения `GET`.

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### <a name="when-to-call-the-method"></a>Когда вызывается данный метод

Если надстройка не может работать без входа пользователя в Office, то при запуске надстройки необходимо вызвать `getAccessTokenAsync` **.

Если некоторые функциональные возможности надстройки не требуют входа пользователя в систему, то при выполнении пользователем операции, требующей входа в систему, можно вызвать `getAccessTokenAsync` **. Значительного снижения производительности при совершении `getAccessTokenAsync` избыточных вызовов не наблюдается, так как Office кэширует и повторно использует маркер доступа до истечения его срока действия, не обращаясь к конечной точке AAD версии 2.0 при каждом вызове `getAccessTokenAsync`. В связи с этим, вызовы `getAccessTokenAsync` можно добавлять во все функции и обработчики, которые инициируют действие, требующее наличия маркера.

### <a name="add-server-side-code"></a>Добавление кода со стороны сервера

Если маркер доступа не передается надстройкой использующему его серверу, то, в большинстве случаев, получение этого маркера особого смысла не имеет. Некоторые задачи, которые могут выполняться надстройкой со стороны сервера, включают в себя:

* Создание одного или нескольких методов веб-API, использующих извлекаемую из маркера информацию о пользователе, например, метода, выполняющего поиск пользовательских настройек в размещенной базе данных (см. **Использование маркера единого доступа в качестве удостоверения** ниже). В зависимости от применяемого языка и платформы, могут оказаться доступными библиотеки, упрощающие написание кода.
* Получение данных Microsoft Graph. Код со стороны сервера должен выполнять следующие операции:

    * Проверка маркеров доступа (см. приведенную ниже статью **Проверка маркера доступа**).
    * Запуск потока «от имени пользователя» с вызовом конечной точки Azure AD версии 2.0, которая включает в себя маркер доступа, некоторые метаданные о пользователе и учетные данные надстройки (ее идентификатор и секрет). В подобной ситуации маркер доступа именуется маркером начальной загрузки.
    * Кэширование нового маркера доступа, возвращаемого потоком, запущенным «от имени пользователя»
    * Получите данные с Microsoft Graph, используя новый маркер.

 Для ознакомления с дополнительной информацией о получении авторизованного доступа к данным пользователя Microsoft Graph см. статью [Авторизованный доступ в Microsoft Graph из вашей надстройки Office](authorize-to-microsoft-graph.md).

#### <a name="validate-the-access-token"></a>Утверждение маркера доступа

После получения веб-API маркера доступа, ему необходимо проверить его до начала использования. Этот маркер представляет собой веб-маркер JSON (JWT), в связи с чем, его проверка осуществляется так же, как проверка маркеров в большинстве стандартных потоков OAuth. Существует несколько библиотек, способных обработать проверку JWT, но основные действия, производимые в ходе ее выполнения, включают в себя:

- проверку правильности формата маркера;
- проверку факта выдачи маркера нужным центром сертификации;
- проверку предназначения маркера для веб-API.

При проверке маркера следует руководствоваться следующими рекомендациями.

- Допустимые маркеры единого входа будет выдаваться центром предоставления полномочий Azure, `https://login.microsoftonline.com`. Содержащееся в маркере утверждение `iss` должно начинаться с этого значения.
- Параметру `aud` маркера будет присваиваться идентификатор приложения, полученный при регистрации надстройки.
- Для параметра `scp` маркера будет задано значение `access_as_user`.

#### <a name="using-the-sso-token-as-an-identity"></a>Использование маркера единого входа в качестве удостоверения

Если надстройке требуется проверить идентификатор пользователя, то информация, которая может использоваться в качестве удостоверения, содержится в маркере единого входа. С идентификатором связаны следующие утверждения, имеющиеся в маркере.

- `name` — отображаемое имя пользователя.
- `preferred_username` — адрес электронной почты пользователя.
- `oid` — GUID, предоставляющий ИД пользователя в Azure Active Directory.
- `tid` — GUID, предоставляющий ИД организации пользователя в Azure Active Directory.

В связи с тем,что значения `name` и `preferred_username` могут меняться, рекомендуется использовать значения `oid` и `tid`, позволяющие коррелировать идентификатор с внутренней службой авторизации.

Так, к примеру, ваша служба может отформатировать все эти значения, как `{oid-value}@{tid-value}`, а затем сохранить их в качестве записи пользователя во внутренней базе данных пользователей. После этого, при подаче последующих запросов, информацию о пользователе можно получить с применением того же самого значения, а доступ к определенным ресурсам может предоставляться с помощью существующих механизмов управления доступом.

### <a name="example-access-token"></a>Пример маркера доступа

Ниже приводятся типовые декодированные данные маркера доступа. Чтобы больше узнать о свойствах, см. [Руководство по маркерам Azure Active Directory версии 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).


```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",         
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",         
    iat: 1521143967,         
    nbf: 1521143967,         
    exp: 1521147867,         
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",         
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",         
    azpacr: "0",         
    e_exp: 262800,         
    name: "Mila Nikolova",         
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",         
    preferred_username: "milan@contoso.com",         
    scp: "access_as_user",         
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",         
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",         
    uti: "MICAQyhrH02ov54bCtIDAA",         
    ver: "2.0"
}
```

## <a name="using-sso-with-an-outlook-add-in"></a>Использование единого доступа с надстройкой Outlook

Есть несколько небольших, но важных отличий в использовании SSO в надстройке Outlook от его использования в надстройке Excel, PowerPoint или Word. Обязательно прочитайте [«Аутентификация пользователя с единым маркером входа» в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) и [сценарии: выполните единый вход в свою службу в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

## <a name="sso-api-reference"></a>Ссылка на API SSO

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

Пространство имен Office Auth `Office.context.auth` предоставляет метод `getAccessTokenAsync`, позволяющий ведущему приложению Office получить маркер доступа к веб-приложению надстройки. Также он косвенно позволяет надстройке получить доступ к данным Microsoft Graph вошедшего в систему пользователя без повторного ввода им учетных данных.

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

Этот метод вызывает конечную точку Azure Active Directory версии 2.0, чтобы получить маркер доступа для веб-приложения надстройки. Это позволяет надстройкам идентифицировать пользователей. Код со стороны сервера может использовать этот маркер, для предоставления веб-приложению надстройки доступа к Microsoft Graph с помощью [потока OAuth, запускаемого «от имени пользователя»](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> В Outlook этот API не поддерживается, если надстройка загружается в почтовый ящик Outlook.com или Gmail.

<table><tr><td>Основные приложения</td><td>Excel, OneNote, Outlook, PowerPoint, Word</td></tr>

 <tr><td>[Наборы требований](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)</td></tr></table>

#### <a name="parameters"></a>Параметры

`options` — не обязательный. Принимает объект `AuthOptions` (см. ниже) для определения метода входа.

`callback` — не обязательный. Принимает метод обратного вызова, способный проанализировать маркер на наличие идентификатора пользователя или использовать его в потоке, запускаемом «от имени пользователя», для получения доступа к Microsoft Graph. Если [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` выполняется успешно, то `AsyncResult.value` представляет собой неформатированный маркер доступа AAD версии 2.0.

После получения Office маркера доступа для надстройки от AAD версии 2.0 с применением метода `getAccessTokenAsync` интерфейс `AuthOptions` предоставляет средства для взаимодействия с пользователем.

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has 
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor 
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development 
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" 
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should 
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```



