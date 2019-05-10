---
title: Элемент Supertip в файле манифеста
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 269a3723db6f98cdb25c61e5a88608c5fb5f3191
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659658"
---
# <a name="supertip"></a><span data-ttu-id="9d5ec-102">Supertip</span><span class="sxs-lookup"><span data-stu-id="9d5ec-102">Supertip</span></span>

<span data-ttu-id="9d5ec-p101">Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Button](control.md#button-control) или [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="9d5ec-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="9d5ec-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9d5ec-105">Child elements</span></span>

|  <span data-ttu-id="9d5ec-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="9d5ec-106">Element</span></span> |  <span data-ttu-id="9d5ec-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9d5ec-107">Required</span></span>  |  <span data-ttu-id="9d5ec-108">Описание</span><span class="sxs-lookup"><span data-stu-id="9d5ec-108">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="9d5ec-109">Title</span><span class="sxs-lookup"><span data-stu-id="9d5ec-109">Title</span></span>](#title) | <span data-ttu-id="9d5ec-110">Да</span><span class="sxs-lookup"><span data-stu-id="9d5ec-110">Yes</span></span> | <span data-ttu-id="9d5ec-111">Текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="9d5ec-111">The text for the supertip.</span></span> |
| [<span data-ttu-id="9d5ec-112">Description</span><span class="sxs-lookup"><span data-stu-id="9d5ec-112">Description</span></span>](#description) | <span data-ttu-id="9d5ec-113">Да</span><span class="sxs-lookup"><span data-stu-id="9d5ec-113">Yes</span></span> | <span data-ttu-id="9d5ec-114">Описание подсказки.</span><span class="sxs-lookup"><span data-stu-id="9d5ec-114">The description for the supertip.</span></span><br><span data-ttu-id="9d5ec-115">**Note**: (Outlook) поддерживаются только клиенты Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="9d5ec-115">**Note**: (Outlook) Only Windows and Mac clients are supported.</span></span> |

### <a name="title"></a><span data-ttu-id="9d5ec-116">Название</span><span class="sxs-lookup"><span data-stu-id="9d5ec-116">Title</span></span>

<span data-ttu-id="9d5ec-p102">Обязательный элемент. Текст суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="9d5ec-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="9d5ec-120">Описание</span><span class="sxs-lookup"><span data-stu-id="9d5ec-120">Description</span></span>

<span data-ttu-id="9d5ec-p103">Обязательный элемент. Описание суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **LongStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="9d5ec-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="9d5ec-124">В Outlook только клиенты Windows и Mac поддерживают элемент **Description** .</span><span class="sxs-lookup"><span data-stu-id="9d5ec-124">For Outlook, only Windows and Mac clients support the **Description** element.</span></span>

## <a name="example"></a><span data-ttu-id="9d5ec-125">Пример</span><span class="sxs-lookup"><span data-stu-id="9d5ec-125">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
