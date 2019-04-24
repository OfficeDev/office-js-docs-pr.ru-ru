---
title: Элемент Supertip в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cdbba342fa591ddff3faf94ecd63a4740fb904da
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450543"
---
# <a name="supertip"></a><span data-ttu-id="3c62b-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="3c62b-102">Supertip</span></span>

<span data-ttu-id="3c62b-p101">Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="3c62b-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="3c62b-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="3c62b-105">Child elements</span></span>

|  <span data-ttu-id="3c62b-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="3c62b-106">Element</span></span> |  <span data-ttu-id="3c62b-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="3c62b-107">Required</span></span>  |  <span data-ttu-id="3c62b-108">Описание</span><span class="sxs-lookup"><span data-stu-id="3c62b-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3c62b-109">Title</span><span class="sxs-lookup"><span data-stu-id="3c62b-109">Title</span></span>](#title)        | <span data-ttu-id="3c62b-110">Да</span><span class="sxs-lookup"><span data-stu-id="3c62b-110">Yes</span></span> |   <span data-ttu-id="3c62b-111">Текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="3c62b-111">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="3c62b-112">Description</span><span class="sxs-lookup"><span data-stu-id="3c62b-112">Description</span></span>](#description)  | <span data-ttu-id="3c62b-113">Да</span><span class="sxs-lookup"><span data-stu-id="3c62b-113">Yes</span></span> |  <span data-ttu-id="3c62b-114">Описание подсказки.</span><span class="sxs-lookup"><span data-stu-id="3c62b-114">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="3c62b-115">Заголовок</span><span class="sxs-lookup"><span data-stu-id="3c62b-115">Title</span></span>

<span data-ttu-id="3c62b-p102">Обязательный элемент. Текст суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="3c62b-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="3c62b-119">Описание</span><span class="sxs-lookup"><span data-stu-id="3c62b-119">Description</span></span>

<span data-ttu-id="3c62b-p103">Обязательный элемент. Описание суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **LongStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="3c62b-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="3c62b-123">Пример</span><span class="sxs-lookup"><span data-stu-id="3c62b-123">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
