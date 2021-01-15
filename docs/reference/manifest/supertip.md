---
title: Элемент Supertip в файле манифеста
description: Элемент Supertip определяет rich tooltip (название и описание).
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 5e8b3850d99f6791726b1b2f0545c5fb4b52c554
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771300"
---
# <a name="supertip"></a><span data-ttu-id="c686b-103">Supertip</span><span class="sxs-lookup"><span data-stu-id="c686b-103">Supertip</span></span>

<span data-ttu-id="c686b-p101">Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="c686b-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c686b-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c686b-106">Child elements</span></span>

|  <span data-ttu-id="c686b-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="c686b-107">Element</span></span> |  <span data-ttu-id="c686b-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c686b-108">Required</span></span>  |  <span data-ttu-id="c686b-109">Описание</span><span class="sxs-lookup"><span data-stu-id="c686b-109">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="c686b-110">Title</span><span class="sxs-lookup"><span data-stu-id="c686b-110">Title</span></span>](#title) | <span data-ttu-id="c686b-111">Да</span><span class="sxs-lookup"><span data-stu-id="c686b-111">Yes</span></span> | <span data-ttu-id="c686b-112">Текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="c686b-112">The text for the supertip.</span></span> |
| [<span data-ttu-id="c686b-113">Description</span><span class="sxs-lookup"><span data-stu-id="c686b-113">Description</span></span>](#description) | <span data-ttu-id="c686b-114">Да</span><span class="sxs-lookup"><span data-stu-id="c686b-114">Yes</span></span> | <span data-ttu-id="c686b-115">Описание подсказки.</span><span class="sxs-lookup"><span data-stu-id="c686b-115">The description for the supertip.</span></span><br><span data-ttu-id="c686b-116">**Примечание.**(Outlook) поддерживаются только клиенты Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="c686b-116">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="c686b-117">Title</span><span class="sxs-lookup"><span data-stu-id="c686b-117">Title</span></span>

<span data-ttu-id="c686b-118">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="c686b-118">Required.</span></span> <span data-ttu-id="c686b-119">Текст суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="c686b-119">The text for the supertip.</span></span> <span data-ttu-id="c686b-120">Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе ShortStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="c686b-120">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="c686b-121">Описание</span><span class="sxs-lookup"><span data-stu-id="c686b-121">Description</span></span>

<span data-ttu-id="c686b-122">Обязательный.</span><span class="sxs-lookup"><span data-stu-id="c686b-122">Required.</span></span> <span data-ttu-id="c686b-123">Описание суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="c686b-123">The description for the supertip.</span></span> <span data-ttu-id="c686b-124">Атрибут **resid** не может быть больше 32 символов и должен иметь значение атрибута **id** элемента **String** в **элементе LongStrings** в [элементе Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="c686b-124">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="c686b-125">В Outlook элемент Description поддерживается только **клиентами** Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="c686b-125">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="c686b-126">Пример</span><span class="sxs-lookup"><span data-stu-id="c686b-126">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
