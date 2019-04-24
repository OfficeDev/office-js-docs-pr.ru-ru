---
title: Элемент Namespace в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: faf77fe8b6bddc734f1b47eb544ffe7e1e7c4aaa
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452104"
---
# <a name="namespace-element"></a><span data-ttu-id="7fdd5-102">Элемент Namespace</span><span class="sxs-lookup"><span data-stu-id="7fdd5-102">Namespace element</span></span>

<span data-ttu-id="7fdd5-103">Определяет пространство имен, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="7fdd5-103">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="7fdd5-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7fdd5-104">Attributes</span></span>

|  <span data-ttu-id="7fdd5-105">Атрибут</span><span class="sxs-lookup"><span data-stu-id="7fdd5-105">Attribute</span></span>  |  <span data-ttu-id="7fdd5-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="7fdd5-106">Required</span></span>  |  <span data-ttu-id="7fdd5-107">Описание</span><span class="sxs-lookup"><span data-stu-id="7fdd5-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="7fdd5-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="7fdd5-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="7fdd5-109">Да</span><span class="sxs-lookup"><span data-stu-id="7fdd5-109">Yes</span></span>  | <span data-ttu-id="7fdd5-110">Должен соответствовать заголовку ShortStrings для пользовательской функции, указанной в элементе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="7fdd5-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="7fdd5-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="7fdd5-111">Child elements</span></span>

<span data-ttu-id="7fdd5-112">Нет</span><span class="sxs-lookup"><span data-stu-id="7fdd5-112">None</span></span>

## <a name="example"></a><span data-ttu-id="7fdd5-113">Пример</span><span class="sxs-lookup"><span data-stu-id="7fdd5-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
