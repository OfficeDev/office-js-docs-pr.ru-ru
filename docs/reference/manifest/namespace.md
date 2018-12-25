---
title: Элемент Namespace в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 8000ea5774b38dd038888c686a33127a2d5bc482
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432328"
---
# <a name="namespace-element"></a><span data-ttu-id="13880-102">Элемент Namespace</span><span class="sxs-lookup"><span data-stu-id="13880-102">Namespace element</span></span>

<span data-ttu-id="13880-103">Определяет пространство имен, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="13880-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="13880-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="13880-104">Attributes</span></span>

|  <span data-ttu-id="13880-105">Атрибут</span><span class="sxs-lookup"><span data-stu-id="13880-105">Attribute</span></span>  |  <span data-ttu-id="13880-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="13880-106">Required</span></span>  |  <span data-ttu-id="13880-107">Описание</span><span class="sxs-lookup"><span data-stu-id="13880-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="13880-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="13880-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="13880-109">Да</span><span class="sxs-lookup"><span data-stu-id="13880-109">Yes</span></span>  | <span data-ttu-id="13880-110">Должен соответствовать заголовку ShortStrings для пользовательской функции, указанной в элементе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="13880-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="13880-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="13880-111">Child elements</span></span>

<span data-ttu-id="13880-112">Нет</span><span class="sxs-lookup"><span data-stu-id="13880-112">None</span></span>

## <a name="example"></a><span data-ttu-id="13880-113">Пример</span><span class="sxs-lookup"><span data-stu-id="13880-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
