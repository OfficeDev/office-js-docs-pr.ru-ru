---
title: Проверка маркера удостоверения надстройки Outlook
description: Надстройка Outlook может отправить вам маркер удостоверения пользователя Exchange, но прежде чем обращаться с запросом как с доверенным, нужно проверить, поступил ли маркер с ожидаемого сервера Exchange Server.
ms.date: 05/08/2020
localization_priority: Normal
ms.openlocfilehash: b416353b0d9875a2024ca4706152472c7e5012b0
ms.sourcegitcommit: 7e6faf3dc144400a7b7e5a42adecbbec0bd4602d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/09/2020
ms.locfileid: "44180212"
---
# <a name="validate-an-exchange-identity-token"></a>Проверка маркера удостоверения Exchange

Надстройка Outlook может отправить вам маркер удостоверения пользователя Exchange, но прежде чем обращаться с запросом как с доверенным, нужно проверить, поступил ли маркер с ожидаемого сервера Exchange Server. Маркеры удостоверений пользователей Exchange представляют собой маркеры JSON Web Token (JWT). Инструкции по проверке JWT представлены в документе [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).

Рекомендуем использовать процесс, состоящий из четырех этапов, для проверки маркера удостоверения и получения уникального идентификатора пользователя. Первый этап: извлечение веб-маркера JSON (JWT) из строки, закодированной в формате URL-адреса Base64. Второй этап: проверка правильности маркера, то есть его предназначения для вашей надстройки Outlook, его актуальности и возможности извлечения допустимого URL-адреса для документа метаданных проверки подлинности. Затем необходимо получить документ метаданных проверки подлинности с сервера Exchange и проверить подпись, приложенную к маркеру удостоверения. Наконец, вычислите уникальный идентификатор для пользователя, объединяя идентификатор Exchange пользователя с URL-адресом документа с метаданными проверки подлинности.

## <a name="extract-the-json-web-token"></a>Извлечение маркера JSON Web Token

Маркер, возвращаемый методом [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods), — это закодированная строка, представляющая его. В этом формате (согласно стандарту RFC 7519) все маркеры JWT состоят из трех частей, разделенных точками. Используется приведенный ниже формат.

```json
{header}.{payload}.{signature}
```

Чтобы получить представление каждой части в формате JSON, необходимо раскодировать заголовок и полезные данные согласно кодировке Base64. Подпись необходимо расшифровать согласно кодировке Base64, чтобы получить массив байтов, содержащий двоичную подпись.

Дополнительные сведения о содержимом маркера см. в статье [Подробные сведения о маркере удостоверения Exchange](inside-the-identity-token.md).

После получения трех раскодированных компонентов можно продолжать проверку содержимого маркера.

## <a name="validate-token-contents"></a>Проверка содержимого маркера

Для проверки содержимого маркера необходимо выполнить указанные ниже действия.

- Проверьте заголовок и убедитесь, что:
    - `typ`для `JWT`утверждения задано значение.
    - `alg`для `RS256`утверждения задано значение.
    - `x5t`присутствует утверждение.

- Проверьте полезную нагрузку и убедитесь, что:
    - `amurl`в параметре `appctx` Claims in the указывается расположение подписанного файла манифеста подписанного маркера. Например, ожидаемое `amurl` значение для Office 365: https://outlook.office365.com:443/autodiscover/metadata/json/1. В следующем разделе [проверяйте домен](#verify-the-domain) на наличие дополнительной информации.
    - Текущее время находится в промежутке между значениями, `nbf` указанными в утверждениях и `exp` . В утверждении `nbf` указано время, с которого начинается срок действия маркера, а в утверждении `exp` — время его окончания. Рекомендуем допускать небольшие различия в заданном времени на разных серверах.
    - `aud`Claim — это ожидаемый URL-адрес надстройки.
    - `version`для `ExIdTok.V1`утверждения в `appctx` утверждении задано значение.

### <a name="verify-the-domain"></a>Проверка домена

При реализации логики проверки, описанной ранее в этом разделе, необходимо также указать, чтобы домен `amurl` утверждения соответствовал домену автообнаружения для пользователя. Для этого необходимо использовать или реализовать службу автообнаружения. Чтобы узнать больше, можно запустить службу [автообнаружения для Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).

## <a name="validate-the-identity-token-signature"></a>Проверка подписи маркера удостоверения

Когда вы убедитесь, что JWT содержит необходимые утверждения, можно переходить к проверке подписи маркера.

### <a name="retrieve-the-public-signing-key"></a>Получение открытого ключа подписывания

Первый этап — получение открытого ключа, соответствующего сертификату, который сервер Exchange Server использовал для подписывания маркера. Этот ключ указан в документе с метаданными проверки подлинности. Этот документ представляет собой JSON-файл, размещенный по URL-адресу, указанному в утверждении `amurl`.

Документ с метаданными проверки подлинности имеет приведенный ниже формат.

```json
{
    "id": "_70b34511-d105-4e2b-9675-39f53305bb01",
    "version": "1.0",
    "name": "Exchange",
    "realm": "*",
    "serviceName": "00000002-0000-0ff1-ce00-000000000000",
    "issuer": "00000002-0000-0ff1-ce00-000000000000@*",
    "allowedAudiences": [
        "00000002-0000-0ff1-ce00-000000000000@*"
    ],
    "keys": [
        {
            "usage": "signing",
            "keyinfo": {
                "x5t": "enh9BJrVPU5ijV1qjZjV-fL2bco"
            },
            "keyvalue": {
                "type": "x509Certificate",
                "value": "MIIHNTCC..."
            }
        }
    ],
    "endpoints": [
        {
            "location": "https://by2pr06mb2229.namprd06.prod.outlook.com:444/autodiscover/metadata/json/1",
            "protocol": "OAuth2",
            "usage": "metadata"
        }
    ]
}
```

Доступные ключи подписывания находятся в массиве `keys`. Выберите подходящий ключ, убедившись, что значение `x5t` в свойстве `keyinfo` совпадает со значением `x5t` в заголовке маркера. Открытый ключ находится в дочернем свойстве `value` свойства `keyvalue`, хранящемся в массиве байтов с кодировкой Base64.

После получения правильного открытого ключа проверьте подпись. Подписанные данные представляют собой первые две части закодированного маркера, разделенные точкой:

```json
{header}.{payload}
```

## <a name="compute-the-unique-id-for-an-exchange-account"></a>Вычисление уникального идентификатора для учетной записи Exchange

Вы можете создать уникальный идентификатор для учетной записи Exchange, объединяя URL-адрес документа метаданных проверки подлинности с идентификатором Exchange для учетной записи. Получив этот уникальный идентификатор, вы можете создать систему единого входа для веб-службы надстройки Outlook. Дополнительные сведения об использовании уникального идентификатора для единого входа см. в статье [Проверка подлинности пользователя с помощью маркера удостоверения для Exchange](authenticate-a-user-with-an-identity-token.md).

## <a name="use-a-library-to-validate-the-token"></a>Проверка маркера с помощью библиотеки

Существует ряд библиотек, способных выполнять общие задачи анализа и проверки JWT. Корпорация Майкрософт предоставляет `System.IdentityModel.Tokens.Jwt` библиотеку, которая может использоваться для проверки маркеров удостоверения пользователей Exchange.

> [!IMPORTANT]
> Мы не рекомендуем использовать управляемый API веб-служб Exchange, так как файл Microsoft. Exchange. WebServices. auth. dll, несмотря на то, что он все еще доступен, теперь устарел и использует неподдерживаемые библиотеки, такие как Microsoft. IdentityModel. Extensions. dll.

### <a name="systemidentitymodeltokensjwt"></a>System.IdentityModel.Tokens.Jwt

Библиотека [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) может анализировать маркер, а также выполнять проверку, но вам потребуется самостоятельно проанализировать утверждение `appctx` и получить открытый ключ подписывания.

```cs
// Load the encoded token
string encodedToken = "...";
JwtSecurityToken jwt = new JwtSecurityToken(encodedToken);

// Parse the appctx claim to get the auth metadata url
string authMetadataUrl = string.Empty;
var appctx = jwt.Claims.FirstOrDefault(claim => claim.Type == "appctx");
if (appctx != null)
{
    var AppContext = JsonConvert.DeserializeObject<ExchangeAppContext>(appctx.Value);

    // Token version check
    if (string.Compare(AppContext.Version, "ExIdTok.V1", StringComparison.InvariantCulture) != 0) {
        // Fail validation
    }

    authMetadataUrl = AppContext.MetadataUrl;
}

// Use System.IdentityModel.Tokens.Jwt library to validate standard parts
JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
TokenValidationParameters tvp = new TokenValidationParameters();

tvp.ValidateIssuer = false;
tvp.ValidateAudience = true;
tvp.ValidAudience = "{URL to add-in}";
tvp.ValidateIssuerSigningKey = true;
// GetSigningKeys downloads the auth metadata doc and
// returns a List<SecurityKey>
tvp.IssuerSigningKeys = GetSigningKeys(authMetadataUrl);
tvp.ValidateLifetime = true;

try
{
    var claimsPrincipal = tokenHandler.ValidateToken(encodedToken, tvp, out SecurityToken validatedToken);

    // If no exception, all standard checks passed
}
catch (SecurityTokenValidationException ex)
{
    // Validation failed
}
```

<br/>

Класс `ExchangeAppContext` определяется следующим образом:

```cs
using Newtonsoft.Json;

/// <summary>
/// Representation of the appctx claim in an Exchange user identity token.
/// </summary>
public class ExchangeAppContext
{
    /// <summary>
    /// The Exchange identifier for the user
    /// </summary>
    [JsonProperty("msexchuid")]
    public string ExchangeUid { get; set; }

    /// <summary>
    /// The token version
    /// </summary>
    public string Version { get; set; }

    /// <summary>
    /// The URL to download authentication metadata
    /// </summary>
    [JsonProperty("amurl")]
    public string MetadataUrl { get; set; }
}
```

Пример проверки маркеров Exchange с помощью этой библиотеки, в котором также реализован метод `GetSigningKeys`: [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).

## <a name="see-also"></a>См. также

- [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
