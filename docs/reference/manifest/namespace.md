---
title: Элемент Namespace в файле манифеста
description: Элемент namespace определяет пространство имен, используемое пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eabd73d3be98271c81723787dd3d1bdb6ee2ebcd
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978671"
---
# <a name="namespace-element"></a><span data-ttu-id="c6bac-103">Элемент Namespace</span><span class="sxs-lookup"><span data-stu-id="c6bac-103">Namespace element</span></span>

<span data-ttu-id="c6bac-104">Определяет пространство имен, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="c6bac-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="c6bac-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c6bac-105">Attributes</span></span>

|  <span data-ttu-id="c6bac-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="c6bac-106">Attribute</span></span>  |  <span data-ttu-id="c6bac-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c6bac-107">Required</span></span>  |  <span data-ttu-id="c6bac-108">Описание</span><span class="sxs-lookup"><span data-stu-id="c6bac-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c6bac-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="c6bac-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="c6bac-110">Нет</span><span class="sxs-lookup"><span data-stu-id="c6bac-110">No</span></span>  | <span data-ttu-id="c6bac-111">Должен соответствовать заголовку ShortStrings для пользовательской функции, указанной в элементе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="c6bac-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="c6bac-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c6bac-112">Child elements</span></span>

<span data-ttu-id="c6bac-113">Нет</span><span class="sxs-lookup"><span data-stu-id="c6bac-113">None</span></span>

## <a name="example"></a><span data-ttu-id="c6bac-114">Пример</span><span class="sxs-lookup"><span data-stu-id="c6bac-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
