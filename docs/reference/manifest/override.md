---
title: Элемент Override в файле манифеста
description: Элемент override позволяет указать значение параметра для дополнительного языкового стандарта.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: aa5d023169389670d15e36f8bee4445529d84711
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611507"
---
# <a name="override-element"></a><span data-ttu-id="0527b-103">Элемент Override</span><span class="sxs-lookup"><span data-stu-id="0527b-103">Override element</span></span>

<span data-ttu-id="0527b-104">Предоставляет способ указать значение параметра для дополнительного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="0527b-104">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="0527b-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="0527b-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0527b-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="0527b-106">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="0527b-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="0527b-107">Contained in</span></span>

|<span data-ttu-id="0527b-108">**Element**</span><span class="sxs-lookup"><span data-stu-id="0527b-108">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="0527b-109">CitationText</span><span class="sxs-lookup"><span data-stu-id="0527b-109">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="0527b-110">Описание</span><span class="sxs-lookup"><span data-stu-id="0527b-110">Description</span></span>](description.md)|
|[<span data-ttu-id="0527b-111">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="0527b-111">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="0527b-112">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="0527b-112">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="0527b-113">DisplayName</span><span class="sxs-lookup"><span data-stu-id="0527b-113">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="0527b-114">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="0527b-114">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="0527b-115">IconUrl</span><span class="sxs-lookup"><span data-stu-id="0527b-115">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="0527b-116">QueryUri</span><span class="sxs-lookup"><span data-stu-id="0527b-116">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="0527b-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0527b-117">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="0527b-118">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="0527b-118">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="0527b-119">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0527b-119">Attributes</span></span>

|<span data-ttu-id="0527b-120">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="0527b-120">**Attribute**</span></span>|<span data-ttu-id="0527b-121">**Тип**</span><span class="sxs-lookup"><span data-stu-id="0527b-121">**Type**</span></span>|<span data-ttu-id="0527b-122">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="0527b-122">**Required**</span></span>|<span data-ttu-id="0527b-123">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0527b-123">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="0527b-124">Языковой стандарт</span><span class="sxs-lookup"><span data-stu-id="0527b-124">Locale</span></span>|<span data-ttu-id="0527b-125">string</span><span class="sxs-lookup"><span data-stu-id="0527b-125">string</span></span>|<span data-ttu-id="0527b-126">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0527b-126">required</span></span>|<span data-ttu-id="0527b-127">Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="0527b-127">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="0527b-128">Значение</span><span class="sxs-lookup"><span data-stu-id="0527b-128">Value</span></span>|<span data-ttu-id="0527b-129">string</span><span class="sxs-lookup"><span data-stu-id="0527b-129">string</span></span>|<span data-ttu-id="0527b-130">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0527b-130">required</span></span>|<span data-ttu-id="0527b-131">Задает значение параметра, представленное для указанного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="0527b-131">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="0527b-132">См. также</span><span class="sxs-lookup"><span data-stu-id="0527b-132">See also</span></span>

- [<span data-ttu-id="0527b-133">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="0527b-133">Localization for Office Add-ins</span></span>](../../develop/localization.md)
