---
ms.date: 01/14/2020
description: Узнайте, как реализовать переменные настраиваемые функции потоковой и автономной работы.
title: Пересчитываемые значения в функциях
localization_priority: Normal
ms.openlocfilehash: a318c87cc5b5f45bf3b1f5fe1341b7008f5a3d2f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609792"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="6d04f-103">Пересчитываемые значения в функциях</span><span class="sxs-lookup"><span data-stu-id="6d04f-103">Volatile values in functions</span></span>

<span data-ttu-id="6d04f-104">Функции volatile — это функции, в которых значение изменяется каждый раз при вычислении ячейки.</span><span class="sxs-lookup"><span data-stu-id="6d04f-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="6d04f-105">Значение может измениться, даже если ни один из аргументов функции не изменится.</span><span class="sxs-lookup"><span data-stu-id="6d04f-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="6d04f-106">Эти функции пересчитываются при каждом пересчете в Excel.</span><span class="sxs-lookup"><span data-stu-id="6d04f-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="6d04f-107">К примеру, представьте себе ячейку, вызывающую функцию `NOW`.</span><span class="sxs-lookup"><span data-stu-id="6d04f-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="6d04f-108">При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.</span><span class="sxs-lookup"><span data-stu-id="6d04f-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="6d04f-109">В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="6d04f-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="6d04f-110">Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="6d04f-110">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="6d04f-111">Пользовательские функции позволяют создавать собственные переменные функции, которые могут быть удобны при обработке дат, времени, случайных чисел и моделирования.</span><span class="sxs-lookup"><span data-stu-id="6d04f-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="6d04f-112">Например, для определения оптимального решения для [имитации Монте Карло](https://en.wikipedia.org/wiki/Monte_Carlo_method) требуется создание случайных входных данных.</span><span class="sxs-lookup"><span data-stu-id="6d04f-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="6d04f-113">При выборе автоматического создания JSON файла объявите переменную с помощью тега Жсдок Comment `@volatile` .</span><span class="sxs-lookup"><span data-stu-id="6d04f-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDoc comment tag `@volatile`.</span></span> <span data-ttu-id="6d04f-114">Дополнительные сведения об автоформировании приведены в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="6d04f-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="6d04f-115">Ниже приведен пример временного настраиваемой функции, которая имитирует пошаговое описание шести костей.</span><span class="sxs-lookup"><span data-stu-id="6d04f-115">An example of a volatile custom function follows, which simulates rolling a six-sided dice.</span></span>

![GIF-файл, в котором показана пользовательская функция, возвращающая случайное значение для имитации шести двусторонних костей](../images/six-sided-die.gif)

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

## <a name="next-steps"></a><span data-ttu-id="6d04f-117">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="6d04f-117">Next steps</span></span>
* <span data-ttu-id="6d04f-118">Сведения о [параметрах настраиваемых функций](custom-functions-parameter-options.md).</span><span class="sxs-lookup"><span data-stu-id="6d04f-118">Learn about [custom functions parameter options](custom-functions-parameter-options.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6d04f-119">См. также</span><span class="sxs-lookup"><span data-stu-id="6d04f-119">See also</span></span>

* [<span data-ttu-id="6d04f-120">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="6d04f-120">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="6d04f-121">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="6d04f-121">Create custom functions in Excel</span></span>](custom-functions-overview.md)
