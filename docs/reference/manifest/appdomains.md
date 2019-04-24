---
title: Элемент AppDomains в файле манифеста
description: ''
ms.date: 12/13/2018
localization_priority: Normal
ms.openlocfilehash: 65391c9529e7ddaa9726d0b58accf90c5b9babef
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450648"
---
# <a name="appdomains-element"></a><span data-ttu-id="cbe81-102">Элемент AppDomains</span><span class="sxs-lookup"><span data-stu-id="cbe81-102">AppDomains element</span></span>

<span data-ttu-id="cbe81-p101">Определяет все домены, кроме указанного в элементе SourceLocation, которые надстройка Office будет использовать для загрузки страниц. Для каждого дополнительного домена укажите элемент AppDomain.</span><span class="sxs-lookup"><span data-stu-id="cbe81-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="cbe81-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="cbe81-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cbe81-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="cbe81-106">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="cbe81-107">Значение каждого элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="cbe81-107">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="cbe81-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="cbe81-108">Contained in</span></span>

[<span data-ttu-id="cbe81-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="cbe81-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="cbe81-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="cbe81-110">Can contain</span></span>

[<span data-ttu-id="cbe81-111">AppDomain</span><span class="sxs-lookup"><span data-stu-id="cbe81-111">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="cbe81-112">Примечания</span><span class="sxs-lookup"><span data-stu-id="cbe81-112">Remarks</span></span>

<span data-ttu-id="cbe81-113">По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="cbe81-113">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="cbe81-114">Для загрузки страниц из других доменов, укажите их домены в элементах **AppDomains** и **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="cbe81-114">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="cbe81-115">Этот элемент не может быть пустым.</span><span class="sxs-lookup"><span data-stu-id="cbe81-115">This element can't be empty.</span></span>
