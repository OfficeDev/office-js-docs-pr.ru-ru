---
title: Элемент AppDomain в файле манифеста
description: Задает дополнительные домены, которые загружают страницы в окне надстройки.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 6990f759df806f24b1d617c036bc1a452e6da38f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718456"
---
# <a name="appdomain-element"></a><span data-ttu-id="2322c-103">Элемент AppDomain</span><span class="sxs-lookup"><span data-stu-id="2322c-103">AppDomain element</span></span>

<span data-ttu-id="2322c-104">Задает дополнительные домены, которые загружают страницы в окне надстройки.</span><span class="sxs-lookup"><span data-stu-id="2322c-104">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="2322c-105">Кроме того, выводит список доверенных доменов, из которых можно создавать вызовы API Office. js из IFrame в надстройке.</span><span class="sxs-lookup"><span data-stu-id="2322c-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="2322c-106">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="2322c-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2322c-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="2322c-107">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="2322c-108">Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="2322c-108">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="2322c-109">*Не* ставьте закрывающую косую черту ("/") для значения.</span><span class="sxs-lookup"><span data-stu-id="2322c-109">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="2322c-110">Содержится в</span><span class="sxs-lookup"><span data-stu-id="2322c-110">Contained in</span></span>

[<span data-ttu-id="2322c-111">AppDomains</span><span class="sxs-lookup"><span data-stu-id="2322c-111">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="2322c-112">Примечания</span><span class="sxs-lookup"><span data-stu-id="2322c-112">Remarks</span></span>

<span data-ttu-id="2322c-113">Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="2322c-113">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="2322c-114">Дополнительные сведения см. в статье [XML-манифест надстроек Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="2322c-114">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
