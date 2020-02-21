---
title: Подробные сведения о маркере удостоверения Exchange в надстройке Outlook
description: Узнайте, из чего состоит маркер удостоверения пользователя Exchange, созданный в надстройке Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4cbbcdc587495a9b490f300414235cba1c5c570a
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166757"
---
# <a name="inside-the-exchange-identity-token"></a><span data-ttu-id="9b310-103">Подробные сведения о маркере удостоверения Exchange</span><span class="sxs-lookup"><span data-stu-id="9b310-103">Inside the Exchange identity token</span></span>

<span data-ttu-id="9b310-104">Маркер удостоверения пользователя Exchange, возвращаемый методом [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods), позволяет надстройке включать удостоверение пользователя в вызовы внутренней службы.</span><span class="sxs-lookup"><span data-stu-id="9b310-104">The Exchange user identity token returned by the [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method provides a way for your add-in code to include the user's identity with calls to your back-end service.</span></span> <span data-ttu-id="9b310-105">В этой статье рассматриваются формат и содержимое маркера.</span><span class="sxs-lookup"><span data-stu-id="9b310-105">This article will discuss the format and contents of the token.</span></span>

<span data-ttu-id="9b310-106">Маркер удостоверения пользователя Exchange — это строка URL-адреса с кодировкой Base64, подписанная сервером Exchange Server, который отправил ее.</span><span class="sxs-lookup"><span data-stu-id="9b310-106">An Exchange user identity token is a base-64 URL-encoded string that is signed by the Exchange server that sent it.</span></span> <span data-ttu-id="9b310-107">Маркер не шифруется, а открытый ключ, используемый для проверки подписи, хранится на сервере Exchange Server, который выдал маркер.</span><span class="sxs-lookup"><span data-stu-id="9b310-107">The token is not encrypted, and the public key that you use to validate the signature is stored on the Exchange server that issued the token.</span></span> <span data-ttu-id="9b310-108">Маркер состоит из трех частей: заголовка, полезных данных и подписи.</span><span class="sxs-lookup"><span data-stu-id="9b310-108">The token has three parts: a header, a payload, and a signature.</span></span> <span data-ttu-id="9b310-109">В строке маркера части отделяются друг от друга точкой (`.`), чтобы маркер было проще разделить.</span><span class="sxs-lookup"><span data-stu-id="9b310-109">In the token string, the parts are separated by a period character (`.`) to make it easy for you to split the token.</span></span>

<span data-ttu-id="9b310-110">В Exchange для маркера удостоверения используется формат JSON Web Token (JWT).</span><span class="sxs-lookup"><span data-stu-id="9b310-110">Exchange uses a the JSON Web Token (JWT) format for the identity token.</span></span> <span data-ttu-id="9b310-111">Сведения о маркерах JWT см. в документе [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span><span class="sxs-lookup"><span data-stu-id="9b310-111">For information about JWT tokens, see [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

## <a name="identity-token-header"></a><span data-ttu-id="9b310-112">Заголовок маркера удостоверения</span><span class="sxs-lookup"><span data-stu-id="9b310-112">Identity token header</span></span>

<span data-ttu-id="9b310-113">Заголовок содержит сведения о формате и подписи маркера.</span><span class="sxs-lookup"><span data-stu-id="9b310-113">The header provides information about the format and signature information of the token.</span></span> <span data-ttu-id="9b310-114">В приведенном ниже примере показано, как выглядит заголовок маркера.</span><span class="sxs-lookup"><span data-stu-id="9b310-114">The following example shows what the header of the token looks like.</span></span>

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
<span data-ttu-id="9b310-115">В приведенной ниже таблице описаны части заголовка маркера.</span><span class="sxs-lookup"><span data-stu-id="9b310-115">The following table describes the parts of the token header.</span></span>

| <span data-ttu-id="9b310-116">Утверждение</span><span class="sxs-lookup"><span data-stu-id="9b310-116">Claim</span></span> | <span data-ttu-id="9b310-117">Значение</span><span class="sxs-lookup"><span data-stu-id="9b310-117">Value</span></span> | <span data-ttu-id="9b310-118">Описание</span><span class="sxs-lookup"><span data-stu-id="9b310-118">Description</span></span> |
|:-----|:-----|:-----|
| `typ` | `JWT` | <span data-ttu-id="9b310-119">Определяет маркер как JSON Web Token.</span><span class="sxs-lookup"><span data-stu-id="9b310-119">Identifies the token as a JSON Web Token.</span></span> <span data-ttu-id="9b310-120">Все маркеры удостоверений, предоставленные сервером Exchange Server, являются маркерами JWT.</span><span class="sxs-lookup"><span data-stu-id="9b310-120">All identity tokens provided by Exchange server are JWT tokens.</span></span> |
| `alg` | `RS256` | <span data-ttu-id="9b310-121">Алгоритм хэширования, используемый для создания подписи.</span><span class="sxs-lookup"><span data-stu-id="9b310-121">The hashing algorithm that is used to create the signature.</span></span> <span data-ttu-id="9b310-122">Все маркеры, предоставляемые сервером Exchange Server, используют алгоритм хэширования RSASSA-PKCS1-v1_5 с SHA-256.</span><span class="sxs-lookup"><span data-stu-id="9b310-122">All tokens provided by Exchange server use the RSASSA-PKCS1-v1_5 with SHA-256 hash algorithm.</span></span> |
| `x5t` | <span data-ttu-id="9b310-123">Отпечаток сертификата</span><span class="sxs-lookup"><span data-stu-id="9b310-123">Certificate thumbprint</span></span> | <span data-ttu-id="9b310-124">Отпечаток маркера X.509.</span><span class="sxs-lookup"><span data-stu-id="9b310-124">The X.509 thumbprint of the token.</span></span> |

## <a name="identity-token-payload"></a><span data-ttu-id="9b310-125">Полезные данные маркера удостоверения</span><span class="sxs-lookup"><span data-stu-id="9b310-125">Identity token payload</span></span>

<span data-ttu-id="9b310-p107">Полезные данные содержат утверждения проверки подлинности, которые идентифицируют учетную запись электронной почты, а также сервер Exchange, который отправляет маркер. В следующем примере показано, как выглядит раздел полезных данных.</span><span class="sxs-lookup"><span data-stu-id="9b310-p107">The payload contains the authentication claims that identify the email account and identify the Exchange server that sent the token. The following example shows what the payload section looks like.</span></span>

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
 
<span data-ttu-id="9b310-128">В приведенной ниже таблице описаны части полезных данных маркера удостоверения.</span><span class="sxs-lookup"><span data-stu-id="9b310-128">The following table lists the parts of the identity token payload.</span></span>

| <span data-ttu-id="9b310-129">Утверждение</span><span class="sxs-lookup"><span data-stu-id="9b310-129">Claim</span></span> | <span data-ttu-id="9b310-130">Описание</span><span class="sxs-lookup"><span data-stu-id="9b310-130">Description</span></span> |
|:-----|:-----|
| `aud` | <span data-ttu-id="9b310-131">URL-адрес надстройки, запросившей маркер.</span><span class="sxs-lookup"><span data-stu-id="9b310-131">The URL of the add-in that requested the token.</span></span> <span data-ttu-id="9b310-132">Маркер действителен, только если он отправлен из надстройки, работающей в браузере клиента.</span><span class="sxs-lookup"><span data-stu-id="9b310-132">A token is only valid if it is sent from the add-in that is running in the client's browser.</span></span> <span data-ttu-id="9b310-133">Если надстройка использует схему манифестов надстроек Office версии 1.1, то этот URL-адрес указан в первом элементе `SourceLocation` под типом формы `ItemRead` или `ItemEdit` (в зависимости от того, какой из них указан первым в элементе [FormSettings](../reference/manifest/formsettings.md) манифеста надстройки).</span><span class="sxs-lookup"><span data-stu-id="9b310-133">If the add-in uses the Office Add-ins manifests schema v1.1, this URL is the URL specified in the first `SourceLocation` element, under the form type `ItemRead` or `ItemEdit`, whichever occurs first as part of the [FormSettings](../reference/manifest/formsettings.md) element in the add-in manifest.</span></span> |
| `iss` | <span data-ttu-id="9b310-p109">Уникальный идентификатор сервера Exchange, выпустившего маркер. Все маркеры, выпущенные сервером Exchange, будут иметь одинаковый идентификатор.</span><span class="sxs-lookup"><span data-stu-id="9b310-p109">A unique identifier for the Exchange server that issued the token. All tokens issued by this Exchange server will have the same identifier.</span></span> |
| `nbf` | <span data-ttu-id="9b310-p110">Дата и время начала срока действия маркера. Значением является количество секунд с 1 января 1970 г.</span><span class="sxs-lookup"><span data-stu-id="9b310-p110">The date and time that the token is valid starting from. The value is the number of seconds since January 1, 1970.</span></span> |
| `exp` | <span data-ttu-id="9b310-p111">Дата и время окончания срока действия маркера. Значением является количество секунд с 1 января 1970 г.</span><span class="sxs-lookup"><span data-stu-id="9b310-p111">The date and time that the token is valid until. The value is the number of seconds since January 1, 1970.</span></span> |
| `appctxsender` | <span data-ttu-id="9b310-140">Уникальный идентификатор для сервера Exchange Server, который отправляет контекст приложения.</span><span class="sxs-lookup"><span data-stu-id="9b310-140">A unique identifier for the Exchange server that sent the application context.</span></span> |
| `isbrowserhostedapp` | <span data-ttu-id="9b310-141">Указывает, размещается ли надстройка в браузере.</span><span class="sxs-lookup"><span data-stu-id="9b310-141">Indicates whether the add-in is hosted in a browser.</span></span> |
| `appctx` | <span data-ttu-id="9b310-142">Контекст приложения для маркера.</span><span class="sxs-lookup"><span data-stu-id="9b310-142">The application context for the token.</span></span> |

<span data-ttu-id="9b310-143">Сведения из утверждения appctx содержат уникальный идентификатор учетной записи и расположение открытого ключа, используемого для подписывания маркера.</span><span class="sxs-lookup"><span data-stu-id="9b310-143">The information in the appctx claim provides you with the unique identifier for the account and the location of the public key used to sign the token.</span></span> <span data-ttu-id="9b310-144">В приведенной ниже таблице перечислены части утверждения `appctx`.</span><span class="sxs-lookup"><span data-stu-id="9b310-144">The following table lists the parts of the `appctx` claim.</span></span>

| <span data-ttu-id="9b310-145">Свойство контекста приложения</span><span class="sxs-lookup"><span data-stu-id="9b310-145">Application context property</span></span> | <span data-ttu-id="9b310-146">Описание</span><span class="sxs-lookup"><span data-stu-id="9b310-146">Description</span></span> |
|:-----|:-----|
| `msexchuid` | <span data-ttu-id="9b310-147">Уникальный идентификатор, связанный с учетной записью электронной почты и сервером Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="9b310-147">A unique identifier associated with the email account and the Exchange server.</span></span> |
| `version` | <span data-ttu-id="9b310-148">Номер версии маркера.</span><span class="sxs-lookup"><span data-stu-id="9b310-148">The version number of the token.</span></span> <span data-ttu-id="9b310-149">Для всех маркеров, предоставленных средой Exchange, используется значение `ExIdTok.V1`.</span><span class="sxs-lookup"><span data-stu-id="9b310-149">For all tokens provided by Exchange, the value is `ExIdTok.V1`.</span></span> |
| `amurl` | <span data-ttu-id="9b310-150">URL-адрес документа метаданных проверки подлинности, который содержит открытый ключ сертификата X.509, который использовался для подписи маркера.</span><span class="sxs-lookup"><span data-stu-id="9b310-150">The URL of the authentication metadata document that contains the public key of the X.509 certificate that was used to sign the token.</span></span><br/><br/><span data-ttu-id="9b310-151">Дополнительные сведения об использовании документа метаданных проверки подлинности см. в статье [Проверка маркера удостоверения Exchange](validate-an-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="9b310-151">For more information about how to use the authentication metadata document, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span> |

## <a name="identity-token-signature"></a><span data-ttu-id="9b310-152">Подпись маркера удостоверения</span><span class="sxs-lookup"><span data-stu-id="9b310-152">Identity token signature</span></span>

<span data-ttu-id="9b310-p114">Подпись создается путем хэширования разделов заголовка и полезных данных с использованием алгоритма, указанного в заголовке, а также самозаверяющего сертификата X509, размещенного на сервере в месте, указанном в полезных данных. Веб-служба может проверить эту подпись, чтобы убедиться в происхождении маркера удостоверения именно на том сервере, который должен был отправить такой маркер.</span><span class="sxs-lookup"><span data-stu-id="9b310-p114">The signature is created by hashing the header and payload sections with the algorithm specified in the header and using the self-signed X509 certificate located on the server at the location specified in the payload. Your web service can validate this signature to help make sure that the identity token comes from the server that you expect to send it.</span></span>

## <a name="see-also"></a><span data-ttu-id="9b310-155">См. также</span><span class="sxs-lookup"><span data-stu-id="9b310-155">See also</span></span>

<span data-ttu-id="9b310-156">Пример, в котором анализируется маркер удостоверения пользователя Exchange: [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span><span class="sxs-lookup"><span data-stu-id="9b310-156">For an example that parses the Exchange user identity token, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>
