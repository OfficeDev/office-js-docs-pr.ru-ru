---
title: Элемент Supertip в файле манифеста
description: Элемент SuperTip определяет расширенную подсказку (название и описание).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: cf88473b72979c839e5d55f44938fda19be24084
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720353"
---
# <a name="supertip"></a><span data-ttu-id="778d4-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="778d4-103">Supertip</span></span>

<span data-ttu-id="778d4-p101">Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="778d4-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="778d4-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="778d4-106">Child elements</span></span>

|  <span data-ttu-id="778d4-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="778d4-107">Element</span></span> |  <span data-ttu-id="778d4-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="778d4-108">Required</span></span>  |  <span data-ttu-id="778d4-109">Описание</span><span class="sxs-lookup"><span data-stu-id="778d4-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="778d4-110">Title</span><span class="sxs-lookup"><span data-stu-id="778d4-110">Title</span></span>](#title) | <span data-ttu-id="778d4-111">Да</span><span class="sxs-lookup"><span data-stu-id="778d4-111">Yes</span></span> | <span data-ttu-id="778d4-112">Текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="778d4-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="778d4-113">Description</span><span class="sxs-lookup"><span data-stu-id="778d4-113">Description</span></span>](#description) | <span data-ttu-id="778d4-114">Да</span><span class="sxs-lookup"><span data-stu-id="778d4-114">Yes</span></span> | <span data-ttu-id="778d4-115">Описание подсказки.</span><span class="sxs-lookup"><span data-stu-id="778d4-115">The description for the supertip.</span></span><br><span data-ttu-id="778d4-116">**Note**: (Outlook) поддерживаются только клиенты Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="778d4-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="778d4-117">Название</span><span class="sxs-lookup"><span data-stu-id="778d4-117">Title</span></span>

<span data-ttu-id="778d4-118">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="778d4-118">Required.</span></span> <span data-ttu-id="778d4-119">Текст суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="778d4-119">The text for the supertip.</span></span> <span data-ttu-id="778d4-120">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="778d4-120">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="778d4-121">Описание</span><span class="sxs-lookup"><span data-stu-id="778d4-121">Description</span></span>

<span data-ttu-id="778d4-122">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="778d4-122">Required.</span></span> <span data-ttu-id="778d4-123">Описание суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="778d4-123">The description for the supertip.</span></span> <span data-ttu-id="778d4-124">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **LongStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="778d4-124">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="778d4-125">В Outlook только клиенты Windows и Mac поддерживают элемент **Description** .</span><span class="sxs-lookup"><span data-stu-id="778d4-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="778d4-126">Пример</span><span class="sxs-lookup"><span data-stu-id="778d4-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
