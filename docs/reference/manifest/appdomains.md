---
title: Элемент AppDomains в файле манифеста
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: cc2f5ade0bdda214c85490f8e474b42f921edbe8
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433683"
---
# <a name="appdomains-element"></a><span data-ttu-id="59bf0-102">Элемент AppDomains</span><span class="sxs-lookup"><span data-stu-id="59bf0-102">AppDomains element</span></span>

<span data-ttu-id="59bf0-p101">Определяет все домены, кроме указанного в элементе SourceLocation, которые надстройка Office будет использовать для загрузки страниц. Для каждого дополнительного домена укажите элемент AppDomain.</span><span class="sxs-lookup"><span data-stu-id="59bf0-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="59bf0-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="59bf0-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="59bf0-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="59bf0-106">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="59bf0-107">Значение каждого элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="59bf0-107">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="59bf0-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="59bf0-108">Contained in</span></span>

[<span data-ttu-id="59bf0-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="59bf0-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="59bf0-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="59bf0-110">Can contain</span></span>

[<span data-ttu-id="59bf0-111">AppDomain</span><span class="sxs-lookup"><span data-stu-id="59bf0-111">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="59bf0-112">Примечания</span><span class="sxs-lookup"><span data-stu-id="59bf0-112">Remarks</span></span>

<span data-ttu-id="59bf0-113">По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="59bf0-113">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="59bf0-114">Для загрузки страниц из других доменов, укажите их домены в элементах **AppDomains** и **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="59bf0-114">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="59bf0-115">Этот элемент не может быть пустым.</span><span class="sxs-lookup"><span data-stu-id="59bf0-115">This element can't be empty.</span></span>
