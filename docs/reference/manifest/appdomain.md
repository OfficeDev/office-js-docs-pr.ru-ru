---
title: Элемент AppDomain в файле манифеста
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: 8216603c87a7dcafde84d25a82f068c9aa86ed96
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870410"
---
# <a name="appdomain-element"></a><span data-ttu-id="9474d-102">Элемент AppDomain</span><span class="sxs-lookup"><span data-stu-id="9474d-102">AppDomain element</span></span>

<span data-ttu-id="9474d-103">Указывает дополнительный домен, который будет использоваться для загрузки страниц в окне надстройки.</span><span class="sxs-lookup"><span data-stu-id="9474d-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="9474d-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="9474d-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9474d-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="9474d-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="9474d-106">Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="9474d-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="9474d-107">*Не* ставьте закрывающую косую черту (/) на значение.</span><span class="sxs-lookup"><span data-stu-id="9474d-107">Do *not* put a closing slash, "/", on the the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="9474d-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="9474d-108">Contained in</span></span>

[<span data-ttu-id="9474d-109">AppDomains</span><span class="sxs-lookup"><span data-stu-id="9474d-109">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="9474d-110">Примечания</span><span class="sxs-lookup"><span data-stu-id="9474d-110">Remarks</span></span>

<span data-ttu-id="9474d-111">Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="9474d-111">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="9474d-112">Дополнительные сведения см. в статье [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="9474d-112">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
