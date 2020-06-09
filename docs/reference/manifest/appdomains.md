---
title: Элемент AppDomains в файле манифеста
description: Перечисляет все домены в дополнение к домену, указанному в `SourceLocation` элементе, который надстройка Office будет использовать для загрузки страниц.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 9183f1815e97bd8d4ac1a7e2cf72d5547d153f7e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608770"
---
# <a name="appdomains-element"></a><span data-ttu-id="1b9a6-103">Элемент AppDomains</span><span class="sxs-lookup"><span data-stu-id="1b9a6-103">AppDomains element</span></span>

<span data-ttu-id="1b9a6-104">Перечисляет все домены в дополнение к домену, указанному в `SourceLocation` элементе, который надстройка Office будет использовать для загрузки страниц.</span><span class="sxs-lookup"><span data-stu-id="1b9a6-104">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="1b9a6-105">Кроме того, выводит список доверенных доменов, из которых можно создавать вызовы API Office. js из IFrame в надстройке.</span><span class="sxs-lookup"><span data-stu-id="1b9a6-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="1b9a6-106">Для каждого дополнительного домена укажите элемент AppDomain.</span><span class="sxs-lookup"><span data-stu-id="1b9a6-106">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="1b9a6-107">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="1b9a6-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1b9a6-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="1b9a6-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="1b9a6-109">Значение каждого элемента **AppDomain** должно содержать протокол (например, `<AppDomain>https://myappdomain<AppDomain>`).</span><span class="sxs-lookup"><span data-stu-id="1b9a6-109">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="1b9a6-110">Содержится в</span><span class="sxs-lookup"><span data-stu-id="1b9a6-110">Contained in</span></span>

[<span data-ttu-id="1b9a6-111">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="1b9a6-111">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="1b9a6-112">Может содержать</span><span class="sxs-lookup"><span data-stu-id="1b9a6-112">Can contain</span></span>

[<span data-ttu-id="1b9a6-113">AppDomain</span><span class="sxs-lookup"><span data-stu-id="1b9a6-113">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="1b9a6-114">Примечания</span><span class="sxs-lookup"><span data-stu-id="1b9a6-114">Remarks</span></span>

<span data-ttu-id="1b9a6-115">По умолчанию надстройка может загружать страницы из домена, указанного в [элементе SourceLocation](sourcelocation.md).</span><span class="sxs-lookup"><span data-stu-id="1b9a6-115">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="1b9a6-116">Для загрузки страниц из других доменов, укажите их домены в элементах **AppDomains** и **AppDomain**.</span><span class="sxs-lookup"><span data-stu-id="1b9a6-116">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="1b9a6-117">Этот элемент не может быть пустым.</span><span class="sxs-lookup"><span data-stu-id="1b9a6-117">This element can't be empty.</span></span>
