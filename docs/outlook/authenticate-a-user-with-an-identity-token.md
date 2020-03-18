---
title: Проверка подлинности пользователя с помощью маркера удостоверения в надстройке
description: Узнайте, как реализовать единый вход в службе с помощью маркера удостоверения, предоставленного надстройкой Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 575bfbf6522a1952525c4da103ee7d6eb54792d4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720850"
---
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a><span data-ttu-id="12fe1-103">Проверка подлинности пользователя с помощью маркера удостоверения для Exchange</span><span class="sxs-lookup"><span data-stu-id="12fe1-103">Authenticate a user with an identity token for Exchange</span></span>

<span data-ttu-id="12fe1-104">Маркеры удостоверений Exchange позволяют надстройке однозначно определять пользователей.
</span><span class="sxs-lookup"><span data-stu-id="12fe1-104">Exchange user identity tokens provide a way for your add-in to uniquely identify an add-in user.</span></span> <span data-ttu-id="12fe1-105">Определив удостоверение пользователя, вы можете реализовать схему проверки подлинности с единым входом для внутренней службы. Благодаря этому пользователи надстроек Outlook смогут подключаться к вашей службе, не выполняя вход.</span><span class="sxs-lookup"><span data-stu-id="12fe1-105">By establishing the user's identity, you can implement a single sign-on (SSO) authentication scheme for your back-end service that enables customers who are using Outlook add-ins to connect to your service without logging in.</span></span> <span data-ttu-id="12fe1-106">Дополнительные сведения о том, в каких случаях следует использовать такие токены, см. в разделе [Маркер удостоверения пользователя Exchange](authentication.md#exchange-user-identity-token).</span><span class="sxs-lookup"><span data-stu-id="12fe1-106">See [Exchange user identity token](authentication.md#exchange-user-identity-token) for more about when to use this token type.</span></span> <span data-ttu-id="12fe1-107">В этой статье мы рассмотрим простой способ проверки подлинности пользователя во внутренней службе с помощью маркера удостоверения Exchange.</span><span class="sxs-lookup"><span data-stu-id="12fe1-107">In this article, we'll take a look at a simplistic method of using the Exchange identity token to authenticate a user to your back-end.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="12fe1-108">Это лишь простой пример реализации единого входа.</span><span class="sxs-lookup"><span data-stu-id="12fe1-108">This is just a simple example of an SSO implementation.</span></span> <span data-ttu-id="12fe1-109">Как всегда, при работе с удостоверениями и проверкой подлинности необходимо убедиться, что код соответствует требованиям к безопасности в вашей организации.</span><span class="sxs-lookup"><span data-stu-id="12fe1-109">As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.</span></span>

## <a name="send-the-id-token-with-each-request"></a><span data-ttu-id="12fe1-110">Отправка маркера удостоверения с каждым запросом</span><span class="sxs-lookup"><span data-stu-id="12fe1-110">Send the ID token with each request</span></span>

<span data-ttu-id="12fe1-111">Для начала надстройка должна получить маркер удостоверения пользователя Exchange с сервера при помощи метода [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="12fe1-111">The first step is for your add-in to obtain the Exchange user identity token from the server by calling [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span></span> <span data-ttu-id="12fe1-112">Затем надстройка отправляет этот маркер с каждым запросом ко внутренней службе.</span><span class="sxs-lookup"><span data-stu-id="12fe1-112">Then the add-in sends this token with every request it makes to your back-end.</span></span> <span data-ttu-id="12fe1-113">Он может быть включен в заголовок или текст запроса.</span><span class="sxs-lookup"><span data-stu-id="12fe1-113">This could be in a header, or as part of the request body.</span></span>

## <a name="validate-the-token"></a><span data-ttu-id="12fe1-114">Проверка маркера</span><span class="sxs-lookup"><span data-stu-id="12fe1-114">Validate the token</span></span>

<span data-ttu-id="12fe1-115">Внутренняя служба ДОЛЖНА проверить маркер, прежде чем принимать его.</span><span class="sxs-lookup"><span data-stu-id="12fe1-115">The back-end MUST validate the token before accepting it.</span></span> <span data-ttu-id="12fe1-116">Очень важно убедиться, что маркер был выдан сервером Exchange Server пользователя.</span><span class="sxs-lookup"><span data-stu-id="12fe1-116">This is an important step to ensure that the token was issued by the user's Exchange server.</span></span> <span data-ttu-id="12fe1-117">Сведения о проверке маркеров удостоверений Exchange см. в статье [Проверка маркера удостоверения Exchange](validate-an-identity-token.md).</span><span class="sxs-lookup"><span data-stu-id="12fe1-117">For information on validating Exchange user identity tokens, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span>

<span data-ttu-id="12fe1-118">После проверки и раскодирования полезные данные маркера выглядят примерно так:</span><span class="sxs-lookup"><span data-stu-id="12fe1-118">Once validated and decoded, the payload of the token looks something like the following.</span></span>

```json
{ 
    "aud" : "https://mailhost.contoso.com/IdentityTest.html",
    "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com",
    "nbf" : "1505749527",
    "exp" : "1505778327",
    "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
    "isbrowserhostedapp":"true",
    "appctx" : {
        "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389",
        "version" : "ExIdTok.V1",
        "amurl" : "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
    }
}
```

## <a name="map-the-token-to-a-user-in-your-backend"></a><span data-ttu-id="12fe1-119">Сопоставление маркера с пользователем во внутренней службе</span><span class="sxs-lookup"><span data-stu-id="12fe1-119">Map the token to a user in your backend</span></span>

<span data-ttu-id="12fe1-120">Внутренняя служба может определить уникальный ИД пользователя на основе маркера и сопоставить его с пользователем во внутренней системе.</span><span class="sxs-lookup"><span data-stu-id="12fe1-120">Your back-end service can calculate a unique user ID from the token and map it to a user in your internal user system.</span></span> <span data-ttu-id="12fe1-121">Например, если для хранения пользователей используется база данных, вы можете добавить уникальный ИД к записи пользователя в ней.</span><span class="sxs-lookup"><span data-stu-id="12fe1-121">For example, if you use a database to store users, you could add this unique ID to the user's record in your database.</span></span>

### <a name="generate-a-unique-id"></a><span data-ttu-id="12fe1-122">Создание уникального идентификатора</span><span class="sxs-lookup"><span data-stu-id="12fe1-122">Generate a unique ID</span></span>

<span data-ttu-id="12fe1-123">Рекомендуем использовать сочетание свойств `msexchuid` и `amurl`.</span><span class="sxs-lookup"><span data-stu-id="12fe1-123">We recommend that you use a combination of the `msexchuid` and `amurl` properties.</span></span> <span data-ttu-id="12fe1-124">Например, вы можете сцепить эти два значения и создать строку в кодировке Base64.</span><span class="sxs-lookup"><span data-stu-id="12fe1-124">For example, you could concatenate the two values together and generate a base 64-encoded string.</span></span> <span data-ttu-id="12fe1-125">Это значение всегда можно получить из маркера, поэтому вы можете сопоставить маркер удостоверения пользователя Exchange с пользователем в системе.</span><span class="sxs-lookup"><span data-stu-id="12fe1-125">This value can be reliably generated from the token every time, so you can map an Exchange user identity token back to the user in your system.</span></span>

### <a name="check-the-user"></a><span data-ttu-id="12fe1-126">Проверка пользователя</span><span class="sxs-lookup"><span data-stu-id="12fe1-126">Check the user</span></span>

<span data-ttu-id="12fe1-127">Создав уникальный идентификатор, необходимо проверить наличие в системе пользователя с этим ИД.</span><span class="sxs-lookup"><span data-stu-id="12fe1-127">With the unique ID generated, the next step is to check for a user in your system with that associated ID.</span></span>

- <span data-ttu-id="12fe1-128">Если пользователь найден, внутренняя служба рассматривает запрос как прошедший проверку подлинности и разрешает продолжить его выполнение.</span><span class="sxs-lookup"><span data-stu-id="12fe1-128">If the user is found, the back-end treats the request as authenticated, and allows the request to proceed.</span></span>

- <span data-ttu-id="12fe1-129">Если же пользователь не найден, внутренняя служба возвращает ошибку, указывающую на то, что пользователь должен выполнить вход.</span><span class="sxs-lookup"><span data-stu-id="12fe1-129">If the user is not found, then the back-end returns an error indicating that the user needs to sign in.</span></span> <span data-ttu-id="12fe1-130">Затем надстройка предлагает пользователю войти во внутреннюю службу, используя имеющийся способ проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="12fe1-130">The add-in then prompts the user to sign in to the back-end using your existing authentication method.</span></span> <span data-ttu-id="12fe1-131">После проверки подлинности пользователя маркер удостоверения Exchange отправляется вместе с другими данными проверки подлинности.
</span><span class="sxs-lookup"><span data-stu-id="12fe1-131">Once the user is authenticated, the Exchange user identity token is submitted with the user authentication details.</span></span> <span data-ttu-id="12fe1-132">Затем внутренняя служба может добавить уникальный идентификатор к записи пользователя в системе.</span><span class="sxs-lookup"><span data-stu-id="12fe1-132">The back-end can then update the user's record in your system with the unique ID.</span></span>
