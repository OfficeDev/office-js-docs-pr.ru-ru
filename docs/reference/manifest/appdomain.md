---
title: Элемент AppDomain в файле манифеста
description: Указывает дополнительные домены, используемые надстройкой, и которые должны быть доверенными для Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778650"
---
# <a name="appdomain-element"></a><span data-ttu-id="857e9-103">Элемент AppDomain</span><span class="sxs-lookup"><span data-stu-id="857e9-103">AppDomain element</span></span>

<span data-ttu-id="857e9-104">Задает дополнительный домен, который должен быть доверенным для Office, в дополнение к тому, что указано в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="857e9-104">Specifies an additional domain that Office should trust, in addition to the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="857e9-105">Указание домена включает в себя следующие эффекты:</span><span class="sxs-lookup"><span data-stu-id="857e9-105">Specifying a domain has these effects:</span></span>

- <span data-ttu-id="857e9-106">Он позволяет открывать страницы, маршруты и другие ресурсы в домене непосредственно в корневой области задач на настольных платформах Office.</span><span class="sxs-lookup"><span data-stu-id="857e9-106">It enables pages, routes, or other resources in the domain to be opened directly in the root task pane of the add-in on desktop Office platforms.</span></span> <span data-ttu-id="857e9-107">(Указание домена в **домене AppDomain** не требуется для Office в Интернете или открытие ресурса в iframe, а также не требуется для открытия ресурса в диалоговом окне, открываемом с помощью [API диалогового окна](../../develop/dialog-api-in-office-add-ins.md).)</span><span class="sxs-lookup"><span data-stu-id="857e9-107">(Specifying a domain in an **AppDomain** isn't necessary for Office on the web or to open a resource in an IFrame, nor it is necessary for opening a resource in a dialog opened with the [Dialog API](../../develop/dialog-api-in-office-add-ins.md).)</span></span>
- <span data-ttu-id="857e9-108">Он позволяет страницам в домене совершать Office.js вызовы API из IFrames в надстройке.</span><span class="sxs-lookup"><span data-stu-id="857e9-108">It enables pages in the domain to make Office.js API calls from IFrames within the add-in.</span></span>

<span data-ttu-id="857e9-109">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="857e9-109">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="857e9-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="857e9-110">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="857e9-111">Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain.com</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="857e9-111">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain.com</AppDomain>`).</span></span>
> 2. <span data-ttu-id="857e9-112">Если для домена существует явный порт, включите его (например, `<AppDomain>https://myappdomain.com:9999</AppDomain>` ).</span><span class="sxs-lookup"><span data-stu-id="857e9-112">If there is an explicit port for the domain, include it (e.g.,`<AppDomain>https://myappdomain.com:9999</AppDomain>`).</span></span>
> 3. <span data-ttu-id="857e9-113">Если дочерний домен должен быть доверенным, включите его (например, `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ).</span><span class="sxs-lookup"><span data-stu-id="857e9-113">If a subdomain needs to be trusted, include it (e.g.,`<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>`).</span></span> <span data-ttu-id="857e9-114">Дочерний домен `mysubdomain.mydomain.com` и `mydomain.com` разные домены.</span><span class="sxs-lookup"><span data-stu-id="857e9-114">The subdomain `mysubdomain.mydomain.com` and `mydomain.com` are different domains.</span></span> <span data-ttu-id="857e9-115">Если необходимо, чтобы оба были доверенными, они должны находиться в отдельных элементах **AppDomain** .</span><span class="sxs-lookup"><span data-stu-id="857e9-115">If both need to be trusted, then both need to be in separate **AppDomain** elements.</span></span>
> 4. <span data-ttu-id="857e9-116">Перечисление того же домена, что и в [элементе SourceLocation](sourcelocation.md) , не оказывает никакого действия и может привести к некоторому определению.</span><span class="sxs-lookup"><span data-stu-id="857e9-116">Listing the same domain as the one specified in the [SourceLocation element](sourcelocation.md) has no effect and may be misleading.</span></span> <span data-ttu-id="857e9-117">В частности, когда вы разрабатываете `localhost` , вам не нужно создавать элемент **AppDomain** для `localhost` .</span><span class="sxs-lookup"><span data-stu-id="857e9-117">In particular, when you are developing on `localhost`, you don't need to create an **AppDomain** element for `localhost`.</span></span>
> 5. <span data-ttu-id="857e9-118">Не включайте ни один из сегментов URL-адреса за пределами домена.</span><span class="sxs-lookup"><span data-stu-id="857e9-118">Don't include any segments of a URL past the domain.</span></span> <span data-ttu-id="857e9-119">Например, не включайте полный URL-адрес страницы.</span><span class="sxs-lookup"><span data-stu-id="857e9-119">For example, don't include the full URL of a page.</span></span>
> 6. <span data-ttu-id="857e9-120">*Не* ставьте закрывающую косую черту ("/") для значения.</span><span class="sxs-lookup"><span data-stu-id="857e9-120">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="857e9-121">Содержится в</span><span class="sxs-lookup"><span data-stu-id="857e9-121">Contained in</span></span>

[<span data-ttu-id="857e9-122">AppDomains</span><span class="sxs-lookup"><span data-stu-id="857e9-122">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="857e9-123">Замечания</span><span class="sxs-lookup"><span data-stu-id="857e9-123">Remarks</span></span>

<span data-ttu-id="857e9-124">Дополнительные сведения см. в статье [XML-манифест надстроек Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="857e9-124">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
