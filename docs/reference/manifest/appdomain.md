---
title: Элемент AppDomain в файле манифеста
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: 2b55f2c1ea7a2a3dc7dec42c913d74006c0f2e3b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433070"
---
# <a name="appdomain-element"></a><span data-ttu-id="8a210-102">Элемент AppDomain</span><span class="sxs-lookup"><span data-stu-id="8a210-102">AppDomain element</span></span>

<span data-ttu-id="8a210-103">Указывает дополнительный домен, который будет использоваться для загрузки страниц в окне надстройки.</span><span class="sxs-lookup"><span data-stu-id="8a210-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="8a210-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="8a210-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8a210-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="8a210-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="8a210-106">Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="8a210-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="8a210-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="8a210-107">Contained in</span></span>

[<span data-ttu-id="8a210-108">AppDomains</span><span class="sxs-lookup"><span data-stu-id="8a210-108">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="8a210-109">Примечания</span><span class="sxs-lookup"><span data-stu-id="8a210-109">Remarks</span></span>

<span data-ttu-id="8a210-110">Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="8a210-110">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="8a210-111">Дополнительные сведения см. в статье [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="8a210-111">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
