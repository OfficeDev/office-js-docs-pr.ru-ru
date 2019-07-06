---
title: Элемент AppDomains в файле манифеста
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: b6db3d46d004021f25edd5733566544010abb457
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575333"
---
# <a name="appdomains-element"></a><span data-ttu-id="1c02e-102">Элемент AppDomains</span><span class="sxs-lookup"><span data-stu-id="1c02e-102">AppDomains element</span></span>

<span data-ttu-id="1c02e-103">Перечисляет все домены в дополнение к домену, указанному в `SourceLocation` элементе, который надстройка Office будет использовать для загрузки страниц.</span><span class="sxs-lookup"><span data-stu-id="1c02e-103">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="1c02e-104">Кроме того, выводит список доверенных доменов, из которых можно создавать вызовы API Office. js из IFrame в надстройке.</span><span class="sxs-lookup"><span data-stu-id="1c02e-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="1c02e-105">Для каждого дополнительного домена укажите элемент AppDomain.</span><span class="sxs-lookup"><span data-stu-id="1c02e-105">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="1c02e-106">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="1c02e-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1c02e-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="1c02e-107">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="1c02e-108">Значение каждого элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="1c02e-108">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="1c02e-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="1c02e-109">Contained in</span></span>

[<span data-ttu-id="1c02e-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="1c02e-110">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="1c02e-111">Может содержать</span><span class="sxs-lookup"><span data-stu-id="1c02e-111">Can contain</span></span>

[<span data-ttu-id="1c02e-112">AppDomain</span><span class="sxs-lookup"><span data-stu-id="1c02e-112">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="1c02e-113">Примечания</span><span class="sxs-lookup"><span data-stu-id="1c02e-113">Remarks</span></span>

<span data-ttu-id="1c02e-114">По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="1c02e-114">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="1c02e-115">Для загрузки страниц из других доменов, укажите их домены в элементах **AppDomains** и **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="1c02e-115">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="1c02e-116">Этот элемент не может быть пустым.</span><span class="sxs-lookup"><span data-stu-id="1c02e-116">This element can't be empty.</span></span>
