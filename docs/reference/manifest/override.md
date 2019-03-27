---
title: Элемент Override в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 020ae490dacbb9b8c493dc022c23d0ebf311a1b9
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870060"
---
# <a name="override-element"></a><span data-ttu-id="407aa-102">Элемент Override</span><span class="sxs-lookup"><span data-stu-id="407aa-102">Override element</span></span>

<span data-ttu-id="407aa-103">Предоставляет способ указать значение параметра для дополнительного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="407aa-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="407aa-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="407aa-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="407aa-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="407aa-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="407aa-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="407aa-106">Contained in</span></span>

|<span data-ttu-id="407aa-107">**Element**</span><span class="sxs-lookup"><span data-stu-id="407aa-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="407aa-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="407aa-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="407aa-109">Описание</span><span class="sxs-lookup"><span data-stu-id="407aa-109">Description</span></span>](description.md)|
|[<span data-ttu-id="407aa-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="407aa-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="407aa-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="407aa-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="407aa-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="407aa-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="407aa-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="407aa-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="407aa-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="407aa-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="407aa-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="407aa-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="407aa-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="407aa-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="407aa-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="407aa-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="407aa-118">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="407aa-118">Attributes</span></span>

|<span data-ttu-id="407aa-119">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="407aa-119">**Attribute**</span></span>|<span data-ttu-id="407aa-120">**Тип**</span><span class="sxs-lookup"><span data-stu-id="407aa-120">**Type**</span></span>|<span data-ttu-id="407aa-121">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="407aa-121">**Required**</span></span>|<span data-ttu-id="407aa-122">**Описание**</span><span class="sxs-lookup"><span data-stu-id="407aa-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="407aa-123">Языковой стандарт</span><span class="sxs-lookup"><span data-stu-id="407aa-123">Locale</span></span>|<span data-ttu-id="407aa-124">string</span><span class="sxs-lookup"><span data-stu-id="407aa-124">string</span></span>|<span data-ttu-id="407aa-125">Обязательный</span><span class="sxs-lookup"><span data-stu-id="407aa-125">required</span></span>|<span data-ttu-id="407aa-126">Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="407aa-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="407aa-127">Значение</span><span class="sxs-lookup"><span data-stu-id="407aa-127">Value</span></span>|<span data-ttu-id="407aa-128">string</span><span class="sxs-lookup"><span data-stu-id="407aa-128">string</span></span>|<span data-ttu-id="407aa-129">Обязательный</span><span class="sxs-lookup"><span data-stu-id="407aa-129">required</span></span>|<span data-ttu-id="407aa-130">Задает значение параметра, представленное для указанного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="407aa-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="407aa-131">См. также</span><span class="sxs-lookup"><span data-stu-id="407aa-131">See also</span></span>

- [<span data-ttu-id="407aa-132">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="407aa-132">Localization for Office Add-ins</span></span>](/office/dev/add-ins/develop/localization)
    
