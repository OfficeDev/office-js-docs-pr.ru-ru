---
title: Элемент Namespace в файле манифеста
description: Элемент namespace определяет пространство имен, используемое пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 45fd0caa039fdeb885cba4b739750fbd8b642252
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718058"
---
# <a name="namespace-element"></a><span data-ttu-id="dd6c4-103">Элемент Namespace</span><span class="sxs-lookup"><span data-stu-id="dd6c4-103">Namespace element</span></span>

<span data-ttu-id="dd6c4-104">Определяет пространство имен, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="dd6c4-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="dd6c4-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="dd6c4-105">Attributes</span></span>

|  <span data-ttu-id="dd6c4-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="dd6c4-106">Attribute</span></span>  |  <span data-ttu-id="dd6c4-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="dd6c4-107">Required</span></span>  |  <span data-ttu-id="dd6c4-108">Описание</span><span class="sxs-lookup"><span data-stu-id="dd6c4-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="dd6c4-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="dd6c4-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="dd6c4-110">Да</span><span class="sxs-lookup"><span data-stu-id="dd6c4-110">Yes</span></span>  | <span data-ttu-id="dd6c4-111">Должен соответствовать заголовку ShortStrings для пользовательской функции, указанной в элементе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="dd6c4-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="dd6c4-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="dd6c4-112">Child elements</span></span>

<span data-ttu-id="dd6c4-113">Нет</span><span class="sxs-lookup"><span data-stu-id="dd6c4-113">None</span></span>

## <a name="example"></a><span data-ttu-id="dd6c4-114">Пример</span><span class="sxs-lookup"><span data-stu-id="dd6c4-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
