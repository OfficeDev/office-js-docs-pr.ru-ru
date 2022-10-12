---
title: Подробные сведения о маркере удостоверения Exchange в надстройке Outlook
description: Узнайте, из чего состоит маркер удостоверения пользователя Exchange, созданный в надстройке Outlook.
ms.date: 10/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7d586203395521deb14e18a3ae52b01459224b75
ms.sourcegitcommit: 787fbe4d4a5462ff6679ad7fd00748bf07391610
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2022
ms.locfileid: "68546433"
---
# <a name="inside-the-exchange-identity-token"></a>Подробные сведения о маркере удостоверения Exchange

Маркер удостоверения пользователя Exchange, возвращаемый методом [getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods), позволяет надстройке включать удостоверение пользователя в вызовы внутренней службы. В этой статье рассматриваются формат и содержимое маркера.

Маркер удостоверения пользователя Exchange — это строка URL-адреса с кодировкой Base64, подписанная сервером Exchange Server, который отправил ее. Маркер не шифруется, а открытый ключ, используемый для проверки подписи, хранится на сервере Exchange Server, который выдал маркер. Маркер состоит из трех частей: заголовка, полезных данных и подписи. В строке маркера части отделяются друг от друга точкой (`.`), чтобы маркер было проще разделить.

В Exchange для маркера удостоверения используется формат JSON Web Token (JWT). Сведения о маркерах JWT см. в документе [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).

## <a name="identity-token-header"></a>Заголовок маркера удостоверения

Заголовок содержит сведения о формате и подписи маркера. В приведенном ниже примере показано, как выглядит заголовок маркера.

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
В приведенной ниже таблице описаны части заголовка маркера.

| Утверждение | Значение | Описание |
|:-----|:-----|:-----|
| `typ` | `JWT` | Определяет маркер как JSON Web Token. Все маркеры удостоверений, предоставленные сервером Exchange Server, являются маркерами JWT. |
| `alg` | `RS256` | Алгоритм хэширования, используемый для создания подписи. Все маркеры, предоставляемые сервером Exchange Server, используют алгоритм хэширования RSASSA-PKCS1-v1_5 с SHA-256. |
| `x5t` | Отпечаток сертификата | Отпечаток маркера X.509. |

## <a name="identity-token-payload"></a>Полезные данные маркера удостоверения

The payload contains the authentication claims that identify the email account and identify the Exchange server that sent the token. The following example shows what the payload section looks like.

```JSON
{ 
  "aud": "https://mailhost.contoso.com/IdentityTest.html", 
  "iss": "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
  "nbf": "1331579055", 
  "exp": "1331607855", 
  "appctxsender": "00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
  "isbrowserhostedapp": "true",
  "appctx": { 
    "msexchuid": "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com",
    "version": "ExIdTok.V1",
    "amurl": "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
  } 
}
```

<br/>
 
В приведенной ниже таблице описаны части полезных данных маркера удостоверения.

| Утверждение | Описание |
|:-----|:-----|
| `aud` | URL-адрес надстройки, запросившей маркер. Маркер действителен, только если он отправлен из надстройки, работающей в браузере клиента. URL-адрес надстройки указывается в манифесте. Разметка зависит от типа манифеста.</br></br>**XML-манифест:** Если надстройка использует схему манифестов надстроек Office версии 1.1, этот URL-адрес является URL-адресом **\<SourceLocation\>** , указанным в первом элементе, `ItemRead` `ItemEdit`в типе формы или (в зависимости от того, что происходит первым как часть элемента [FormSettings](/javascript/api/manifest/formsettings) в манифесте надстройки).</br></br>**Манифест Teams (предварительная версия):** URL-адрес указывается в свойстве extensions.audienceClaimUrl. |
| `iss` | Уникальный идентификатор сервера Exchange, выпустившего маркер. Все маркеры, выпущенные сервером Exchange, будут иметь одинаковый идентификатор. |
| `nbf` | The date and time that the token is valid starting from. The value is the number of seconds since January 1, 1970. |
| `exp` | The date and time that the token is valid until. The value is the number of seconds since January 1, 1970. |
| `appctxsender` | Уникальный идентификатор для сервера Exchange Server, который отправляет контекст приложения. |
| `isbrowserhostedapp` | Указывает, размещается ли надстройка в браузере. |
| `appctx` | Контекст приложения для маркера. |

Сведения из утверждения appctx содержат уникальный идентификатор учетной записи и расположение открытого ключа, используемого для подписывания маркера. В приведенной ниже таблице перечислены части утверждения `appctx`.

| Свойство контекста приложения | Описание |
|:-----|:-----|
| `msexchuid` | Уникальный идентификатор, связанный с учетной записью электронной почты и сервером Exchange Server. |
| `version` | Номер версии маркера. Для всех маркеров, предоставленных средой Exchange, используется значение `ExIdTok.V1`. |
| `amurl` | URL-адрес документа метаданных проверки подлинности, который содержит открытый ключ сертификата X.509, который использовался для подписи маркера.<br/><br/>Дополнительные сведения об использовании документа метаданных проверки подлинности см. в статье [Проверка маркера удостоверения Exchange](validate-an-identity-token.md). |

## <a name="identity-token-signature"></a>Подпись маркера удостоверения

The signature is created by hashing the header and payload sections with the algorithm specified in the header and using the self-signed X509 certificate located on the server at the location specified in the payload. Your web service can validate this signature to help make sure that the identity token comes from the server that you expect to send it.

## <a name="see-also"></a>См. также

Пример, в котором анализируется маркер удостоверения пользователя Exchange: [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).
