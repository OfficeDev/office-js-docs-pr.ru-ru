---
title: Элемент AppDomains в файле манифеста
description: Список всех доменов в дополнение к домену, указанному в `SourceLocation` элементе, который будет использоваться вашей надстройкой Office и должен быть доверенным для Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778657"
---
# <a name="appdomains-element"></a><span data-ttu-id="c1746-103">Элемент AppDomains</span><span class="sxs-lookup"><span data-stu-id="c1746-103">AppDomains element</span></span>

<span data-ttu-id="c1746-104">Перечисляет все домены в дополнение к домену, указанному в `SourceLocation` элементе, что ваша надстройка Office будет использовать и должна быть доверенной для Office.</span><span class="sxs-lookup"><span data-stu-id="c1746-104">Lists any domains, in addition to the domain specified in the `SourceLocation` element, that your Office Add-in will use and that should be trusted by Office.</span></span> <span data-ttu-id="c1746-105">Это позволяет страницам в доменах совершать вызовы Office.js API из IFrames в надстройке и имеет другие эффекты.</span><span class="sxs-lookup"><span data-stu-id="c1746-105">This enables pages in the domains to make calls to Office.js APIs from IFrames within the add-in and has other effects.</span></span> <span data-ttu-id="c1746-106">Для каждого дополнительного домена укажите элемент **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="c1746-106">For each additional domain, specify an **AppDomain** element.</span></span>

 <span data-ttu-id="c1746-107">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="c1746-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c1746-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="c1746-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="c1746-109">Существуют ограничения на то, что может быть значением элемента **AppDomain** .</span><span class="sxs-lookup"><span data-stu-id="c1746-109">There are restrictions on what can be the value of a **AppDomain** element.</span></span> <span data-ttu-id="c1746-110">Дополнительные сведения см. в разделе [AppDomain](appdomain.md).</span><span class="sxs-lookup"><span data-stu-id="c1746-110">For more information, see [AppDomain](appdomain.md).</span></span>

## <a name="contained-in"></a><span data-ttu-id="c1746-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="c1746-111">Contained in</span></span>

[<span data-ttu-id="c1746-112">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c1746-112">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c1746-113">Может содержать</span><span class="sxs-lookup"><span data-stu-id="c1746-113">Can contain</span></span>

[<span data-ttu-id="c1746-114">AppDomain</span><span class="sxs-lookup"><span data-stu-id="c1746-114">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="c1746-115">Примечания</span><span class="sxs-lookup"><span data-stu-id="c1746-115">Remarks</span></span>

<span data-ttu-id="c1746-116">По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="c1746-116">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="c1746-117">Этот элемент не может быть пустым.</span><span class="sxs-lookup"><span data-stu-id="c1746-117">This element can't be empty.</span></span>
