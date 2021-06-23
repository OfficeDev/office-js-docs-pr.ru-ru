---
ms.date: 01/14/2020
description: Узнайте, как реализовать нестабильную и офлайновую потоковую передачу пользовательских функций.
title: Пересчитываемые значения в функциях
localization_priority: Normal
ms.openlocfilehash: f441ef4fb7f90add5318546e3ccf4cc8bc60a8cf
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075889"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="53a8c-103">Пересчитываемые значения в функциях</span><span class="sxs-lookup"><span data-stu-id="53a8c-103">Volatile values in functions</span></span>

<span data-ttu-id="53a8c-104">Летучие функции — это функции, в которых значение меняется при каждом расчете ячейки.</span><span class="sxs-lookup"><span data-stu-id="53a8c-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="53a8c-105">Значение может измениться, даже если ни один из аргументов функции не изменится.</span><span class="sxs-lookup"><span data-stu-id="53a8c-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="53a8c-106">Эти функции пересчитываются при каждом пересчете в Excel.</span><span class="sxs-lookup"><span data-stu-id="53a8c-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="53a8c-107">К примеру, представьте себе ячейку, вызывающую функцию `NOW`.</span><span class="sxs-lookup"><span data-stu-id="53a8c-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="53a8c-108">При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.</span><span class="sxs-lookup"><span data-stu-id="53a8c-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="53a8c-109">В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="53a8c-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="53a8c-110">Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="53a8c-110">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="53a8c-111">Настраиваемые функции позволяют создавать собственные летучие функции, которые могут быть полезны при обработке дат, времени, случайных чисел и моделирования.</span><span class="sxs-lookup"><span data-stu-id="53a8c-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="53a8c-112">Например, [моделирование Монте-Карло](https://en.wikipedia.org/wiki/Monte_Carlo_method) требует генерации случайных входных данных для определения оптимального решения.</span><span class="sxs-lookup"><span data-stu-id="53a8c-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="53a8c-113">Если вы решили автогенерировать файл JSON, заявите о волатильной функции с помощью тега комментариев JSDoc. `@volatile`</span><span class="sxs-lookup"><span data-stu-id="53a8c-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDoc comment tag `@volatile`.</span></span> <span data-ttu-id="53a8c-114">Дополнительные сведения об автогенерации см. в [метаданных Autogenerate JSON для пользовательских функций.](custom-functions-json-autogeneration.md)</span><span class="sxs-lookup"><span data-stu-id="53a8c-114">From more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="53a8c-115">Пример волатильной настраиваемой функции, которая имитирует развертывание шестистолковой кости.</span><span class="sxs-lookup"><span data-stu-id="53a8c-115">An example of a volatile custom function follows, which simulates rolling a six-sided dice.</span></span>

![GIF показывает настраиваемую функцию, возвращая случайное значение для имитации прокатки шести сторон кости.](../images/six-sided-die.gif)

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided() {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a><span data-ttu-id="53a8c-117">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="53a8c-117">Next steps</span></span>
* <span data-ttu-id="53a8c-118">Узнайте о [настраиваемом параметре функций](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="53a8c-118">Learn about [custom functions parameter options](custom-functions-parameter-options.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="53a8c-119">См. также</span><span class="sxs-lookup"><span data-stu-id="53a8c-119">See also</span></span>

* [<span data-ttu-id="53a8c-120">Вручную создайте метаданные JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="53a8c-120">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="53a8c-121">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="53a8c-121">Create custom functions in Excel</span></span>](custom-functions-overview.md)
