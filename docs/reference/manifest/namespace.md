---
title: Элемент Namespace в файле манифеста
description: Элемент Namespace определяет пространство имен, которое пользовательская функция использует в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 342f5ebcafa861838956f1033f8597cf05e60215
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771263"
---
# <a name="namespace-element"></a><span data-ttu-id="53bce-103">Элемент Namespace</span><span class="sxs-lookup"><span data-stu-id="53bce-103">Namespace element</span></span>

<span data-ttu-id="53bce-104">Определяет пространство имен, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="53bce-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="53bce-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="53bce-105">Attributes</span></span>

|  <span data-ttu-id="53bce-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="53bce-106">Attribute</span></span>  |  <span data-ttu-id="53bce-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="53bce-107">Required</span></span>  |  <span data-ttu-id="53bce-108">Описание</span><span class="sxs-lookup"><span data-stu-id="53bce-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="53bce-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="53bce-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="53bce-110">Нет</span><span class="sxs-lookup"><span data-stu-id="53bce-110">No</span></span>  | <span data-ttu-id="53bce-111">Должен соответствовать заголовку ShortStrings для пользовательской функции, указанной в элементе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="53bce-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> <span data-ttu-id="53bce-112">Может быть не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="53bce-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="53bce-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="53bce-113">Child elements</span></span>

<span data-ttu-id="53bce-114">Нет</span><span class="sxs-lookup"><span data-stu-id="53bce-114">None</span></span>

## <a name="example"></a><span data-ttu-id="53bce-115">Пример</span><span class="sxs-lookup"><span data-stu-id="53bce-115">Example</span></span>

```xml
<Namespace resid="namespace" />
```
