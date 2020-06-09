---
title: Элемент Type в файле манифеста
description: Элемент Type указывает, является ли эквивалентная надстройка надстройкой COM или XLL.
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: b59f903af39facd7543e7384189817d5365cf8c9
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604561"
---
# <a name="type-element"></a><span data-ttu-id="074fe-103">Элемент Type</span><span class="sxs-lookup"><span data-stu-id="074fe-103">Type element</span></span>

<span data-ttu-id="074fe-104">Указывает, является ли эквивалентная надстройка надстройкой COM или XLL.</span><span class="sxs-lookup"><span data-stu-id="074fe-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="074fe-105">**Тип надстройки:** Область задач, настраиваемая функция</span><span class="sxs-lookup"><span data-stu-id="074fe-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="074fe-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="074fe-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="074fe-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="074fe-107">Contained in</span></span>

[<span data-ttu-id="074fe-108">Екуивалентадд</span><span class="sxs-lookup"><span data-stu-id="074fe-108">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="074fe-109">Значения типов надстроек</span><span class="sxs-lookup"><span data-stu-id="074fe-109">Add-in type values</span></span>

<span data-ttu-id="074fe-110">Необходимо указать одно из следующих значений для `Type` элемента.</span><span class="sxs-lookup"><span data-stu-id="074fe-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="074fe-111">COM: эквивалентная надстройка — это надстройка COM.</span><span class="sxs-lookup"><span data-stu-id="074fe-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="074fe-112">XLL: определяет эквивалентную надстройку в формате XLL.</span><span class="sxs-lookup"><span data-stu-id="074fe-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="074fe-113">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="074fe-113">See also</span></span>

- [<span data-ttu-id="074fe-114">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="074fe-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="074fe-115">Обеспечение совместимости надстройки Excel с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="074fe-115">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)