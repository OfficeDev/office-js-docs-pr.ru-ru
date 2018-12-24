---
title: Элемент Supertip в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: bae997eda8e1055c5be76382456ba83acca7b91c
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433672"
---
# <a name="supertip"></a><span data-ttu-id="c9550-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="c9550-102">Supertip</span></span>

<span data-ttu-id="c9550-p101">Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="c9550-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c9550-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c9550-105">Child elements</span></span>

|  <span data-ttu-id="c9550-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="c9550-106">Element</span></span> |  <span data-ttu-id="c9550-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c9550-107">Required</span></span>  |  <span data-ttu-id="c9550-108">Описание</span><span class="sxs-lookup"><span data-stu-id="c9550-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c9550-109">Title</span><span class="sxs-lookup"><span data-stu-id="c9550-109">Title</span></span>](#title)        | <span data-ttu-id="c9550-110">Да</span><span class="sxs-lookup"><span data-stu-id="c9550-110">Yes</span></span> |   <span data-ttu-id="c9550-111">Текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="c9550-111">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="c9550-112">Description</span><span class="sxs-lookup"><span data-stu-id="c9550-112">Description</span></span>](#description)  | <span data-ttu-id="c9550-113">Да</span><span class="sxs-lookup"><span data-stu-id="c9550-113">Yes</span></span> |  <span data-ttu-id="c9550-114">Описание подсказки.</span><span class="sxs-lookup"><span data-stu-id="c9550-114">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="c9550-115">Title</span><span class="sxs-lookup"><span data-stu-id="c9550-115">Title</span></span>

<span data-ttu-id="c9550-p102">Обязательный элемент. Текст суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="c9550-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="c9550-119">Описание</span><span class="sxs-lookup"><span data-stu-id="c9550-119">Description</span></span>

<span data-ttu-id="c9550-p103">Обязательный элемент. Описание суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **LongStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="c9550-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="c9550-123">Пример</span><span class="sxs-lookup"><span data-stu-id="c9550-123">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
