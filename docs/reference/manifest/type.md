---
title: Тип элемента в файле манифеста
description: Элемент Type указывает, является ли эквивалентная надстройка com надстройка или XLL.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 5af3359c232e91b097311bfc06fc9b1c932b0703
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836811"
---
# <a name="type-element"></a><span data-ttu-id="f0aab-103">Элемент Type</span><span class="sxs-lookup"><span data-stu-id="f0aab-103">Type element</span></span>

<span data-ttu-id="f0aab-104">Указывает, является ли эквивалентная надстройка com надстройка или XLL.</span><span class="sxs-lookup"><span data-stu-id="f0aab-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="f0aab-105">**Тип надстройки:** Области задач, настраиваемая функция</span><span class="sxs-lookup"><span data-stu-id="f0aab-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="f0aab-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="f0aab-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="f0aab-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="f0aab-107">Contained in</span></span>

[<span data-ttu-id="f0aab-108">EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="f0aab-108">EquivalentAddin</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="f0aab-109">Значения типа надстройки</span><span class="sxs-lookup"><span data-stu-id="f0aab-109">Add-in type values</span></span>

<span data-ttu-id="f0aab-110">Необходимо указать одно из следующих значений `Type` элемента.</span><span class="sxs-lookup"><span data-stu-id="f0aab-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="f0aab-111">COM. Указывает, что эквивалентная надстройка — это надстройка COM.</span><span class="sxs-lookup"><span data-stu-id="f0aab-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="f0aab-112">XLL. Указывает эквивалентную надстройка Excel XLL.</span><span class="sxs-lookup"><span data-stu-id="f0aab-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="f0aab-113">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="f0aab-113">See also</span></span>

- [<span data-ttu-id="f0aab-114">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="f0aab-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="f0aab-115">Убедитесь, что надстройка Office совместима с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="f0aab-115">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)