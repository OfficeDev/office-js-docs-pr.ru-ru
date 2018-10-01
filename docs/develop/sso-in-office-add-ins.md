---
title: Включение единого входа для надстроек Office
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: 05b5088a61df3f77a09b60dbdc3129074d5f8530
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348172"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Включение единого входа для надстроек Office (предварительная версия)

Пользователи входят в Office (онлайн, на мобильной или настольной платформе), используя личную, рабочую или учебную учетную запись Майкрософт (Office 365). Воспользуйтесь удобной функцией единого входа для однократной авторизации пользователя в своей надстройке без необходимости повторного входа.

![Изображение, иллюстрирующее процесс входа в надстройку](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a>Статус предварительной версии

API единого входа в настоящее время поддерживается только в предварительной версии. Он доступен разработчикам в экспериментальных целях. Его не следует применять в рабочих надстройках. Кроме того, надстройки, в которых используется единый вход, не принимаются в [AppSource](https://appsource.microsoft.com).

Предварительную версию службы единого входа поддерживают не все приложения Office. Она доступна для Word, Excel, Outlook и PowerPoint. Дополнительные сведения о текущей поддержке API единого входа см. в статье [Наборы обязательных элементов API идентификации](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).

### <a name="requirements-and-best-practices"></a>Требования и рекомендации

Чтобы использовать единый вход, необходимо подключить бета-версию библиотеки JavaScript для Office из `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` на начальной HTML-странице надстройки.

Если вы работаете с надстройкой **Outlook**, обязательно включите современную проверку подлинности для клиента Office 365. Сведения о том, как это сделать, см. в статье [Exchange Online: как включить в клиенте современную проверку подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Вы *не* должны полагаться на службу единого входа как на единственный метод проверки подлинности. Необходимо реализовать альтернативную систему проверки подлинности, к которой ваша надстройка сможет обратиться в случае ошибок. Можно использовать систему пользовательских таблиц и проверки подлинности или задействовать одного из поставщиков входа социальных сетей. Дополнительные сведения о том, как это сделать с помощью надстройки Office, см. в статье [авторизация внешних служб в надстройке Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). Для *Outlook* существует рекомендованная альтернативная система. Дополнительные сведения см. в статье [Сценарий: реализация единого входа для службы в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

### <a name="how-sso-works-at-runtime"></a>Принцип работы единого входа во время выполнения

На приведенной ниже схеме показано, как работает единый вход.

![Диаграмма, демонстрирующая процесс единого входа](../images/sso-overview-diagram.png)

1. Код JavaScript надстройки вызывает новый API Office.js — [](#sso-api-reference). Он указывает ведущему приложению Office, что необходимо получить маркер доступа к надстройке См. раздел [Пример маркера доступа](#example-access-token).
2. Если вход в Office не выполнен, в ведущем приложении открывается всплывающее окно, в котором пользователю предлагается войти.
3. Если пользователь запускает надстройку в первый раз, ему предлагается дать согласие.
4. Ведущее приложение Office запрашивает **маркер надстройки** у конечной точки Azure AD версии 2.0 для текущего пользователя.
5. Azure AD отправляет маркер надстройки ведущему приложению Office.
6. Ведущее приложение Office отправляет **маркер** надстройке в составе объекта результата, возвращенного при вызове метода `getAccessTokenAsync`.
7. JavaScript в надстройке может проанализировать маркер и извлечь необходимую информацию, например, адрес электронной почты пользователя. 
8. Кроме того, надстройка может отправить HTTP-запрос на сервер для получения дополнительных сведений о пользователе, например, его настроек. Можно также отправить маркер доступа на сервер для анализа и проверки. 

## <a name="develop-an-sso-add-in"></a>Разработка надстройки с единым входом

В этом разделе описаны задачи, необходимые для создания надстройки Office с единым входом. Эти задачи описываются независимо от языка и платформы. Подробные пошаговые инструкции см. в следующих статьях:

* [Создание надстройки Office на платформе Node.js с использованием единого входа](create-sso-office-add-ins-nodejs.md)
* [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>Создание приложения-службы

Зарегистрируйте надстройку на портале регистрации конечной точки Azure v2.0: https://apps.dev.microsoft.com. Этот процесс занимает 5 – 10 минут и включает выполнение следующих задач:

* Получение идентификатора и секрета клиента для надстройки.
* Укажите разрешения, которые необходимы надстройкам для AAD v. 2.0 (при необходимости — для Microsoft Graph); разрешение "профиля" требуется всегда;
* Предоставление надстройке доверия ведущего приложения Office.
* предварительная авторизация ведущего приложения Office для надстройки с помощью заданного по умолчанию разрешения *access_as_user*.

Для получения дополнительной информации об этом процессе см. статью [Регистрация надстройки Office, использующей единый вход с конечной точкой Azure AD версии 2.0](register-sso-add-in-aad-v2.md).

### <a name="configure-the-add-in"></a>Конфигурация надстройки

Добавьте новую разметку в манифест надстройки:

* **WebApplicationInfo** — родительский элемент для указанных ниже элементов;
* **Id** — идентификатор клиента надстройки; это идентификатор приложения, который вы получаете в рамках регистрации надстройки; См. статью [Регистрация надстройки Office, использующей единый вход с конечной точкой Azure AD версии 2.0](register-sso-add-in-aad-v2.md).
* **Resource** — URL-адрес надстройки;
* **Scopes** — родительский элемент одного или нескольких элементов **Scope**;
* **Область** — указывает разрешение, необходимое надстройке для работы с AAD. Разрешение `profile` требуется всегда, и оно может быть единственным необходимым разрешением, если надстройка не получает доступ к Microsoft Graph. Если надстройка получает этот доступ, потребуются элементы **Scope** для необходимых разрешений Microsoft Graph; например, `User.Read`, `Mail.Read`. Для библиотек, которые используются в коде для доступа к Microsoft Graph, могут потребоваться дополнительные разрешения. Например, для библиотеки проверки подлинности Майкрософт (MSAL) для .NET требуется разрешение `offline_access`. Для получения дополнительной информации см. статью [Авторизованный доступ в Microsoft Graph из вашей надстройки Office](authorize-to-microsoft-graph.md).

Для всех ведущих приложений, кроме Outlook, добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Для Outlook добавьте разметку в конец раздела `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

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
> В данном примере явным образом обрабатывается только один тип ошибки. Для ознакомления с примерами более сложной обработки ошибок см. статьи [Home.js в Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) и [program.js в Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). См. статью [Сообщения устранения ошибок единого входа (SSO)](troubleshoot-sso-in-office-add-ins.md).
 

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

Далее представлен пример передачи маркера надстройки на сервер. При отправке запроса обратно на сервер маркер указывается в качестве заголовка `Authorization`. Данный пример предусматривает отправку данных JSON, поэтому используется метод `POST`, однако `GET` достаточно для отправки маркера доступа, если не выполняется запись в сервер.

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

#### <a name="when-to-call-the-method"></a>Когда вызывать метод

Если надстройка не может работать без входа в Office, необходимо вызвать `getAccessTokenAsync` *при запуске надстройки*.

Если в надстройке присутствует функциональность, которая не требует входа пользователя, метод `getAccessTokenAsync` *вызывается тогда, когда пользователь выполняет действие, для которого требуется вход*. Нет значительного замедления при повторяющихся вызовах `getAccessTokenAsync`, поскольку Office кэширует маркер доступа и использует его снова, пока не истечет срок его действия, не вызывая конечную точку AAD v. 2.0 при каждом вызове  `getAccessTokenAsync`. Поэтому вызовы `getAccessTokenAsync` можно добавлять во все функции и обработчики, которые инициируют действие, где нужен маркер.

### <a name="add-server-side-code"></a>Добавление серверного кода

В большинстве случаев практически нет смысла получать маркер доступа, если надстройка не передает его на сторону сервера и не использует его там. Далее указаны некоторые серверные задачи, которые может выполнять надстройка.

* Создание одного или нескольких методов веб-API, использующих информацию о пользователе, которая извлекается из маркера, например, метод поиска предпочтений пользователя в базе данных на сервере (См. статью **Использование маркера единого входа в качестве удостоверения** далее). В зависимости от языка и платформы могут быть доступны библиотеки, который упростят создание нужного кода.
* Получение данных Microsoft Graph. Серверный код должен:

    * проверять маркеры доступа (см. статью **Проверка маркера доступа** далее);
    * Инициируйте поток «от имени» с вызовом конечной точки Azure AD версии 2.0, который включает токен доступа, некоторые метаданные о пользователе и учетные данные надстройки (ее идентификатор и секрет). в этом контексте маркер доступа называется маркером начальной загрузки;
    * выполнять кэширование нового маркера доступа от потока "от имени";
    * Получите данные с Microsoft Graph, используя новый маркер.

 Для ознакомления с дополнительной информацией о получении авторизованного доступа к данным пользователя Microsoft Graph см. статью [Авторизованный доступ в Microsoft Graph из вашей надстройки Office](authorize-to-microsoft-graph.md).

#### <a name="validate-the-access-token"></a>Утвердите маркер доступа

Когда веб-API получит маркер доступа, этот токен необходимо проверить перед использованием. Это маркер JSON Web Token (JWT), то есть его проверка выполняется так же, как и в большинстве стандартных потоков OAuth. Доступно множество библиотек, которые могут выполнять проверку JWT, но основные действия подразумевают:

- проверку правильности формата маркера;
- проверку факта выдачи маркера нужным центром сертификации;
- проверку предназначения маркера для веб-API.

При проверке маркера следует учитывать приведенные ниже рекомендации.

- Действительные маркеры единого входа выдает центр сертификации Azure, `https://login.microsoftonline.com`. Утверждение `iss` в маркере должно начинаться с этого значения.
- Параметру `aud` маркера будет присвоено значение идентификатора приложения с портала регистрации.
- Для параметра `scp` маркера будет задано значение `access_as_user`.

#### <a name="using-the-sso-token-as-an-identity"></a>Использование маркера единого входа в качестве удостоверения

Если приложению необходимо проверить удостоверение пользователя, то маркер единого входа содержит сведения, с помощью которых можно такое удостоверение определить. Ниже перечислены утверждения из маркера, связанные с удостоверениями.

- `name` — Отображаемое имя пользователя.
- `preferred_username` — Адрес электронной почты пользователя.
- `oid` — GUID, предоставляющий ИД пользователя в Azure Active Directory.
- `tid` — GUID, предоставляющий ИД организации пользователя в Azure Active Directory.

Значения `name` и `preferred_username` могут меняться, поэтому рекомендуется использовать значения `oid` и `tid`, чтобы коррелировать удостоверение с внутренней службой авторизации.

Например, если служба может форматировать эти значения вместе (в виде `{oid-value}@{tid-value}`), то их следует хранить в качестве значения в записи пользователя во внутренней базе данных пользователей. При последующих запросах удостоверение пользователя можно будет получать с помощью того же значения, а доступ к определенным ресурсам может предоставляться в соответствии с действующими механизмами управления доступом.

### <a name="example-access-token"></a>Пример маркера доступа

Далее представлены типичные расшифрованные полезные данные маркера доступа. Для получения дополнительной информации о свойствах см. статью [Ссылка на маркеры Azure Active Directory v2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).


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

## <a name="using-sso-with-an-outlook-add-in"></a>Использование SSO с надстройкой Outlook

Имеются небольшие, но важные различия между использованием функции единого входа в надстройке Outlook и использованием ее в надстройке Excel, PowerPoint или Word. Ознакомьтесь со статьями [Аутентификация пользователя с маркером единого входа в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) и [Сценарий: реализация единого входа для службы в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

## <a name="sso-api-reference"></a>Справка по API SSO

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

Проверка подлинности пространства имен Office `Office.context.auth` предоставляет метод `getAccessTokenAsync`, позволяющий основному приложению Office получать маркер доступа к веб-приложению надстройки. Косвенно это также позволяет надстройке получать доступ к данным Microsoft Graph с включенным пользователем, не требуя от пользователя входа во второй раз.

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

Этот метод вызывает конечную точку Azure Active Directory версии 2.0, чтобы получить токен доступа к веб-приложению надстройки. Это позволяет надстройкам идентифицировать пользователей. Код на стороне сервера может использовать этот маркер для доступа к Microsoft Graph, чтобы добавить веб-приложение надстройки с помощью [потока OAuth "от имени пользователя"](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> В Outlook этот интерфейс API не поддерживается, если надстройка загружается в почтовый ящик Outlook.com или Gmail.

<table><tr><td>Основные приложения</td><td>Excel, OneNote, Outlook, PowerPoint, Word</td></tr>

 <tr><td>Наборы обязательных элементов</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a>Параметры

`options` — Необязательный параметр. Принимает `AuthOptions` объект (см. ниже) для определения расширений функциональности входа.

`callback` — Необязательный параметр. Принимает метод обратного вызова, который может анализировать маркер для идентификатора пользователя или использовать маркер в потоке «от имени», чтобы получить доступ к Microsoft Graph. Если [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` «успешно завершено», тогда `AsyncResult.value` — это необработанный AAD v. отформатированный маркер доступа 2.0.

`AuthOptions` Интерфейс предоставляет параметры для взаимодействия с пользователем, когда Office получает маркер доступа для надстройки от AAD v. 2.0 с методом `getAccessTokenAsync`.

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



