---
title: Элемент Type в файле манифеста
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628230"
---
# <a name="type-element"></a><span data-ttu-id="277d7-102">Элемент Type</span><span class="sxs-lookup"><span data-stu-id="277d7-102">Type element</span></span>

<span data-ttu-id="277d7-103">Указывает, является ли эквивалентная надстройка надстройкой COM или XLL.</span><span class="sxs-lookup"><span data-stu-id="277d7-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="277d7-104">**Тип надстройки:** Область задач, настраиваемая функция</span><span class="sxs-lookup"><span data-stu-id="277d7-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="277d7-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="277d7-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="277d7-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="277d7-106">Contained in</span></span>

[<span data-ttu-id="277d7-107">Екуивалентадд</span><span class="sxs-lookup"><span data-stu-id="277d7-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="277d7-108">Значения типов надстроек</span><span class="sxs-lookup"><span data-stu-id="277d7-108">Add-in type values</span></span>

<span data-ttu-id="277d7-109">Необходимо указать одно из следующих значений для `Type` элемента.</span><span class="sxs-lookup"><span data-stu-id="277d7-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="277d7-110">COM: эквивалентная надстройка — это надстройка COM.</span><span class="sxs-lookup"><span data-stu-id="277d7-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="277d7-111">XLL: определяет эквивалентную надстройку в формате XLL.</span><span class="sxs-lookup"><span data-stu-id="277d7-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="277d7-112">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="277d7-112">See also</span></span>

- [<span data-ttu-id="277d7-113">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="277d7-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="277d7-114">Обеспечение совместимости надстройки Excel с существующей надстройкой COM</span><span class="sxs-lookup"><span data-stu-id="277d7-114">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)