---
ms.date: 05/03/2019
description: Узнайте, как реализовать переменные настраиваемые функции потоковой и автономной работы.
title: Переменные значения в функциях
localization_priority: Normal
ms.openlocfilehash: 1ca3edc3de2d9ac5f2171004f89466352c5cfa1e
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2019
ms.locfileid: "33627999"
---
# <a name="volatile-values-in-functions"></a><span data-ttu-id="e6335-103">Переменные значения в функциях</span><span class="sxs-lookup"><span data-stu-id="e6335-103">Volatile values in functions</span></span>

<span data-ttu-id="e6335-104">Функции volatile — это функции, в которых значение изменяется каждый раз при вычислении ячейки.</span><span class="sxs-lookup"><span data-stu-id="e6335-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="e6335-105">Значение может измениться, даже если ни один из аргументов функции не изменится.</span><span class="sxs-lookup"><span data-stu-id="e6335-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="e6335-106">Эти функции пересчитываются при каждом пересчете в Excel.</span><span class="sxs-lookup"><span data-stu-id="e6335-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="e6335-107">К примеру, представьте себе ячейку, вызывающую функцию `NOW`.</span><span class="sxs-lookup"><span data-stu-id="e6335-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="e6335-108">При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.</span><span class="sxs-lookup"><span data-stu-id="e6335-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="e6335-109">В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="e6335-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="e6335-110">Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="e6335-110">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="e6335-111">Пользовательские функции позволяют создавать собственные переменные функции, которые могут быть удобны при обработке дат, времени, случайных чисел и моделирования.</span><span class="sxs-lookup"><span data-stu-id="e6335-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="e6335-112">Например, для определения оптимального решения для [имитации Монте Карло](https://en.wikipedia.org/wiki/Monte_Carlo_method
) требуется создание случайных входных данных.</span><span class="sxs-lookup"><span data-stu-id="e6335-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method
) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="e6335-113">При выборе автоматического создания JSON файла объявите переменную с помощью тега `@volatile`жсдок Comment.</span><span class="sxs-lookup"><span data-stu-id="e6335-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.</span></span> <span data-ttu-id="e6335-114">Дополнительные сведения об автоформировании приведены в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="e6335-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="e6335-115">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="e6335-115">Next steps</span></span>
<span data-ttu-id="e6335-116">Сведения о том, как [сохранить состояние в пользовательских функциях](custom-functions-save-state.md).</span><span class="sxs-lookup"><span data-stu-id="e6335-116">Learn how to [save state in your custom functions](custom-functions-save-state.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e6335-117">См. также</span><span class="sxs-lookup"><span data-stu-id="e6335-117">See also</span></span>

* [<span data-ttu-id="e6335-118">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="e6335-118">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="e6335-119">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="e6335-119">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="e6335-120">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="e6335-120">Create custom functions in Excel</span></span>](custom-functions-overview.md)
