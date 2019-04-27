---
title: Элемент Type в файле манифеста
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356921"
---
# <a name="type-element"></a><span data-ttu-id="65c00-102">Элемент Type</span><span class="sxs-lookup"><span data-stu-id="65c00-102">Type element</span></span>

<span data-ttu-id="65c00-103">Указывает, является ли эквивалентная надстройка надстройкой COM или XLL.</span><span class="sxs-lookup"><span data-stu-id="65c00-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="65c00-104">**Тип надстройки:** Область задач, настраиваемая функция</span><span class="sxs-lookup"><span data-stu-id="65c00-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="65c00-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="65c00-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="65c00-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="65c00-106">Contained in</span></span>

[<span data-ttu-id="65c00-107">Екуивалентадд</span><span class="sxs-lookup"><span data-stu-id="65c00-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="65c00-108">Значения типов надстроек</span><span class="sxs-lookup"><span data-stu-id="65c00-108">Add-in type values</span></span>

<span data-ttu-id="65c00-109">Необходимо указать одно из следующих значений для `Type` элемента.</span><span class="sxs-lookup"><span data-stu-id="65c00-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="65c00-110">COM: эквивалентная надстройка — это надстройка COM.</span><span class="sxs-lookup"><span data-stu-id="65c00-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="65c00-111">XLL: определяет эквивалентную надстройку в формате XLL.</span><span class="sxs-lookup"><span data-stu-id="65c00-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="65c00-112">См. также</span><span class="sxs-lookup"><span data-stu-id="65c00-112">See also</span></span>

- [<span data-ttu-id="65c00-113">Обеспечение совместимости пользовательских функций с пользовательскими функциями XLL</span><span class="sxs-lookup"><span data-stu-id="65c00-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="65c00-114">Обеспечение совместимости надстройки Office с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="65c00-114">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)