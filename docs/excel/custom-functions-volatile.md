---
ms.date: 04/30/2019
description: Узнайте, как реализовать переменные настраиваемые функции потоковой и автономной работы.
title: Переменные значения в функциях (Предварительная версия)
localization_priority: Normal
ms.openlocfilehash: 63618adecff57398e1630e6b5ab43c0dbc753b36
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527325"
---
## <a name="volatile-values-in-functions"></a><span data-ttu-id="18b42-103">Переменные значения в функциях</span><span class="sxs-lookup"><span data-stu-id="18b42-103">Volatile values in functions</span></span>

<span data-ttu-id="18b42-104">Функции volatile — это функции, в которых значение изменяется каждый раз при вычислении ячейки.</span><span class="sxs-lookup"><span data-stu-id="18b42-104">Volatile functions are functions in which the value changes each time the cell is calculated.</span></span> <span data-ttu-id="18b42-105">Значение может измениться, даже если ни один из аргументов функции не изменится.</span><span class="sxs-lookup"><span data-stu-id="18b42-105">The value can change even if none of the function's arguments change.</span></span> <span data-ttu-id="18b42-106">Эти функции пересчитываются при каждом пересчете в Excel.</span><span class="sxs-lookup"><span data-stu-id="18b42-106">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="18b42-107">К примеру, представьте себе ячейку, вызывающую функцию `NOW`.</span><span class="sxs-lookup"><span data-stu-id="18b42-107">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="18b42-108">При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.</span><span class="sxs-lookup"><span data-stu-id="18b42-108">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="18b42-109">В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="18b42-109">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="18b42-110">Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="18b42-110">For a comprehensive list of Excel’s volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="18b42-111">Пользовательские функции позволяют создавать собственные переменные функции, которые могут быть удобны при обработке дат, времени, случайных чисел и моделирования.</span><span class="sxs-lookup"><span data-stu-id="18b42-111">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modeling.</span></span> <span data-ttu-id="18b42-112">Например, для определения оптимального решения для [имитации Монте Карло](https://en.wikipedia.org/wiki/Monte_Carlo_method
) требуется создание случайных входных данных.</span><span class="sxs-lookup"><span data-stu-id="18b42-112">For example, [Monte Carlo simulations](https://en.wikipedia.org/wiki/Monte_Carlo_method
) require the generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="18b42-113">При выборе автоматического создания JSON файла объявите переменную с помощью тега `@volatile`жсдок Comment.</span><span class="sxs-lookup"><span data-stu-id="18b42-113">If choosing to autogenerate your JSON file, declare a volatile function with the JSDOC comment tag `@volatile`.</span></span> <span data-ttu-id="18b42-114">Дополнительные сведения об автоформировании приведены в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="18b42-114">From more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="18b42-115">См. также</span><span class="sxs-lookup"><span data-stu-id="18b42-115">See also</span></span>

* [<span data-ttu-id="18b42-116">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="18b42-116">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="18b42-117">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="18b42-117">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="18b42-118">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="18b42-118">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="18b42-119">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="18b42-119">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="18b42-120">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="18b42-120">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
