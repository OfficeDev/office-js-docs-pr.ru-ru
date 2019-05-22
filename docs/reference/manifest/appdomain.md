---
title: Элемент AppDomain в файле манифеста
description: ''
ms.date: 05/15/2019
localization_priority: Normal
ms.openlocfilehash: b1d71648cc7646eec246f3d0a8113c843eed2e74
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337197"
---
# <a name="appdomain-element"></a><span data-ttu-id="e8ab8-102">Элемент AppDomain</span><span class="sxs-lookup"><span data-stu-id="e8ab8-102">AppDomain element</span></span>

<span data-ttu-id="e8ab8-103">Указывает дополнительный домен, который будет использоваться для загрузки страниц в окне надстройки.</span><span class="sxs-lookup"><span data-stu-id="e8ab8-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="e8ab8-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="e8ab8-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e8ab8-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="e8ab8-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="e8ab8-106">Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="e8ab8-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="e8ab8-107">*Не* ставьте закрывающую косую черту ("/") для значения.</span><span class="sxs-lookup"><span data-stu-id="e8ab8-107">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="e8ab8-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="e8ab8-108">Contained in</span></span>

[<span data-ttu-id="e8ab8-109">AppDomains</span><span class="sxs-lookup"><span data-stu-id="e8ab8-109">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="e8ab8-110">Примечания</span><span class="sxs-lookup"><span data-stu-id="e8ab8-110">Remarks</span></span>

<span data-ttu-id="e8ab8-111">Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="e8ab8-111">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="e8ab8-112">Дополнительные сведения см. в статье [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests).</span><span class="sxs-lookup"><span data-stu-id="e8ab8-112">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
