---
title: Элемент Override в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1d2400312f12116b1ac5f4010135541e783dcc7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432867"
---
# <a name="override-element"></a><span data-ttu-id="6430d-102">Элемент Override</span><span class="sxs-lookup"><span data-stu-id="6430d-102">Override element</span></span>

<span data-ttu-id="6430d-103">Предоставляет способ указать значение параметра для дополнительного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="6430d-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="6430d-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="6430d-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6430d-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="6430d-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="6430d-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="6430d-106">Contained in</span></span>

|<span data-ttu-id="6430d-107">**Element**</span><span class="sxs-lookup"><span data-stu-id="6430d-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="6430d-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="6430d-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="6430d-109">Description</span><span class="sxs-lookup"><span data-stu-id="6430d-109">Description</span></span>](description.md)|
|[<span data-ttu-id="6430d-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="6430d-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="6430d-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="6430d-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="6430d-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="6430d-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="6430d-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="6430d-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="6430d-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="6430d-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="6430d-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="6430d-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="6430d-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="6430d-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="6430d-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="6430d-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="6430d-118">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6430d-118">Attributes</span></span>

|<span data-ttu-id="6430d-119">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="6430d-119">**Attribute**</span></span>|<span data-ttu-id="6430d-120">**Тип**</span><span class="sxs-lookup"><span data-stu-id="6430d-120">**Type**</span></span>|<span data-ttu-id="6430d-121">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="6430d-121">**Required**</span></span>|<span data-ttu-id="6430d-122">**Описание**</span><span class="sxs-lookup"><span data-stu-id="6430d-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="6430d-123">Locale</span><span class="sxs-lookup"><span data-stu-id="6430d-123">Locale</span></span>|<span data-ttu-id="6430d-124">string</span><span class="sxs-lookup"><span data-stu-id="6430d-124">string</span></span>|<span data-ttu-id="6430d-125">Обязательный</span><span class="sxs-lookup"><span data-stu-id="6430d-125">required</span></span>|<span data-ttu-id="6430d-126">Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.</span><span class="sxs-lookup"><span data-stu-id="6430d-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="6430d-127">Значение</span><span class="sxs-lookup"><span data-stu-id="6430d-127">Value</span></span>|<span data-ttu-id="6430d-128">string</span><span class="sxs-lookup"><span data-stu-id="6430d-128">string</span></span>|<span data-ttu-id="6430d-129">Обязательный</span><span class="sxs-lookup"><span data-stu-id="6430d-129">required</span></span>|<span data-ttu-id="6430d-130">Задает значение параметра, представленное для указанного языкового стандарта.</span><span class="sxs-lookup"><span data-stu-id="6430d-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="6430d-131">См. также</span><span class="sxs-lookup"><span data-stu-id="6430d-131">See also</span></span>

- [<span data-ttu-id="6430d-132">Локализация надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6430d-132">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
