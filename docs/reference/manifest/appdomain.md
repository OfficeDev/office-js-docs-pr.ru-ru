---
title: Элемент AppDomain в файле манифеста
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: da28b3b4dec5d669462a781db3c0628bd32c7182
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596790"
---
# <a name="appdomain-element"></a><span data-ttu-id="5090d-102">Элемент AppDomain</span><span class="sxs-lookup"><span data-stu-id="5090d-102">AppDomain element</span></span>

<span data-ttu-id="5090d-103">Задает дополнительные домены, которые загружают страницы в окне надстройки.</span><span class="sxs-lookup"><span data-stu-id="5090d-103">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="5090d-104">Кроме того, выводит список доверенных доменов, из которых можно создавать вызовы API Office. js из IFrame в надстройке.</span><span class="sxs-lookup"><span data-stu-id="5090d-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="5090d-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="5090d-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="5090d-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="5090d-106">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="5090d-107">Значение элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain</AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="5090d-107">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="5090d-108">*Не* ставьте закрывающую косую черту ("/") для значения.</span><span class="sxs-lookup"><span data-stu-id="5090d-108">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="5090d-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="5090d-109">Contained in</span></span>

[<span data-ttu-id="5090d-110">AppDomains</span><span class="sxs-lookup"><span data-stu-id="5090d-110">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="5090d-111">Примечания</span><span class="sxs-lookup"><span data-stu-id="5090d-111">Remarks</span></span>

<span data-ttu-id="5090d-112">Элементы **AppDomain** следует использовать для указания дополнительных доменов, отличных от указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="5090d-112">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="5090d-113">Дополнительные сведения см. в статье [XML-манифест надстроек Office](../../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="5090d-113">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
