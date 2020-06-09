---
title: Элемент Namespace в файле манифеста
description: Элемент namespace определяет пространство имен, используемое пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f4b3510c6c137bd303af8a3eaac8ebe66c5f4dc7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612236"
---
# <a name="namespace-element"></a><span data-ttu-id="bb4c4-103">Элемент Namespace</span><span class="sxs-lookup"><span data-stu-id="bb4c4-103">Namespace element</span></span>

<span data-ttu-id="bb4c4-104">Определяет пространство имен, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="bb4c4-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="bb4c4-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bb4c4-105">Attributes</span></span>

|  <span data-ttu-id="bb4c4-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="bb4c4-106">Attribute</span></span>  |  <span data-ttu-id="bb4c4-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="bb4c4-107">Required</span></span>  |  <span data-ttu-id="bb4c4-108">Описание</span><span class="sxs-lookup"><span data-stu-id="bb4c4-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bb4c4-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="bb4c4-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="bb4c4-110">Нет</span><span class="sxs-lookup"><span data-stu-id="bb4c4-110">No</span></span>  | <span data-ttu-id="bb4c4-111">Должен соответствовать заголовку ShortStrings для пользовательской функции, указанной в элементе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="bb4c4-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="bb4c4-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="bb4c4-112">Child elements</span></span>

<span data-ttu-id="bb4c4-113">Нет</span><span class="sxs-lookup"><span data-stu-id="bb4c4-113">None</span></span>

## <a name="example"></a><span data-ttu-id="bb4c4-114">Пример</span><span class="sxs-lookup"><span data-stu-id="bb4c4-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
