---
title: Элемент Supertip в файле манифеста
description: Элемент SuperTip определяет расширенную подсказку (название и описание).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 8061c9dcd7903db0f1265084498d6c86654e1dfa
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608721"
---
# <a name="supertip"></a><span data-ttu-id="d562b-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="d562b-103">Supertip</span></span>

<span data-ttu-id="d562b-p101">Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="d562b-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d562b-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="d562b-106">Child elements</span></span>

|  <span data-ttu-id="d562b-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="d562b-107">Element</span></span> |  <span data-ttu-id="d562b-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d562b-108">Required</span></span>  |  <span data-ttu-id="d562b-109">Описание</span><span class="sxs-lookup"><span data-stu-id="d562b-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="d562b-110">Title</span><span class="sxs-lookup"><span data-stu-id="d562b-110">Title</span></span>](#title) | <span data-ttu-id="d562b-111">Да</span><span class="sxs-lookup"><span data-stu-id="d562b-111">Yes</span></span> | <span data-ttu-id="d562b-112">Текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="d562b-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="d562b-113">Description</span><span class="sxs-lookup"><span data-stu-id="d562b-113">Description</span></span>](#description) | <span data-ttu-id="d562b-114">Да</span><span class="sxs-lookup"><span data-stu-id="d562b-114">Yes</span></span> | <span data-ttu-id="d562b-115">Описание подсказки.</span><span class="sxs-lookup"><span data-stu-id="d562b-115">The description for the supertip.</span></span><br><span data-ttu-id="d562b-116">**Note**: (Outlook) поддерживаются только клиенты Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="d562b-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="d562b-117">Название</span><span class="sxs-lookup"><span data-stu-id="d562b-117">Title</span></span>

<span data-ttu-id="d562b-118">Обязательный элемент.</span><span class="sxs-lookup"><span data-stu-id="d562b-118">Required.</span></span> <span data-ttu-id="d562b-119">Текст суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="d562b-119">The text for the supertip.</span></span> <span data-ttu-id="d562b-120">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **ShortStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="d562b-120">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="d562b-121">Описание</span><span class="sxs-lookup"><span data-stu-id="d562b-121">Description</span></span>

<span data-ttu-id="d562b-122">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="d562b-122">Required.</span></span> <span data-ttu-id="d562b-123">Описание суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="d562b-123">The description for the supertip.</span></span> <span data-ttu-id="d562b-124">Атрибуту **Resid** должно быть присвоено значение атрибута **ID** элемента **String** в элементе **LongStrings** элемента [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="d562b-124">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="d562b-125">В Outlook только клиенты Windows и Mac поддерживают элемент **Description** .</span><span class="sxs-lookup"><span data-stu-id="d562b-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="d562b-126">Пример</span><span class="sxs-lookup"><span data-stu-id="d562b-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
