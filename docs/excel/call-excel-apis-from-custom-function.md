---
title: Вызов API Microsoft Excel из настраиваемой функции
description: Узнайте, какие API Microsoft Excel вы можете вызывать из пользовательской функции.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: a25d3f151f648560ee24a3da3f689cb9767bd52a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609806"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a><span data-ttu-id="a5161-103">Вызов API Microsoft Excel из настраиваемой функции</span><span class="sxs-lookup"><span data-stu-id="a5161-103">Call Microsoft Excel APIs from a custom function</span></span>

<span data-ttu-id="a5161-104">Вызовите API Office. js Excel из пользовательских функций, чтобы получить данные диапазона и получить дополнительный контекст для вычислений.</span><span class="sxs-lookup"><span data-stu-id="a5161-104">Call Office.js Excel APIs from your custom functions to get range data and obtain more context for your calculations.</span></span>

<span data-ttu-id="a5161-105">Вызов API Office. js с помощью настраиваемой функции может быть полезен в следующих случаях:</span><span class="sxs-lookup"><span data-stu-id="a5161-105">Calling Office.js APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="a5161-106">Перед вычислением пользовательская функция должна получить сведения из Excel.</span><span class="sxs-lookup"><span data-stu-id="a5161-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="a5161-107">Эти сведения могут включать в себя свойства документов, форматы диапазонов, пользовательские XML-части, имя книги или другие сведения, относящиеся к Excel.</span><span class="sxs-lookup"><span data-stu-id="a5161-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="a5161-108">Настраиваемая функция будет задавать числовой формат ячейки для возвращаемых значений после вычисления.</span><span class="sxs-lookup"><span data-stu-id="a5161-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

## <a name="code-sample"></a><span data-ttu-id="a5161-109">Пример кода</span><span class="sxs-lookup"><span data-stu-id="a5161-109">Code sample</span></span>

<span data-ttu-id="a5161-110">Для вызова API Office. js первым нужен контекст.</span><span class="sxs-lookup"><span data-stu-id="a5161-110">To call into the Office.js APIs you first need a context.</span></span> <span data-ttu-id="a5161-111">Используйте `Excel.RequestContext` объект для получения контекста.</span><span class="sxs-lookup"><span data-stu-id="a5161-111">Use the `Excel.RequestContext` object to get a context.</span></span> <span data-ttu-id="a5161-112">Затем используйте контекст для вызова API, которые необходимы в книге.</span><span class="sxs-lookup"><span data-stu-id="a5161-112">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="a5161-113">В приведенном ниже примере кода показано, как получить диапазон значений из книги.</span><span class="sxs-lookup"><span data-stu-id="a5161-113">The following code sample shows how to get a range of values from the workbook.</span></span>

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a><span data-ttu-id="a5161-114">Ограничения на вызов Office. js с помощью настраиваемой функции</span><span class="sxs-lookup"><span data-stu-id="a5161-114">Limitations of calling Office.js through a custom function</span></span>

<span data-ttu-id="a5161-115">Не вызывайте API Office. js из пользовательской функции, которая изменяет среду Excel.</span><span class="sxs-lookup"><span data-stu-id="a5161-115">Don't call Office.js APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="a5161-116">Это означает, что пользовательские функции не должны выполнять следующие действия:</span><span class="sxs-lookup"><span data-stu-id="a5161-116">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="a5161-117">Вставка, удаление или форматирование ячеек в электронной таблице.</span><span class="sxs-lookup"><span data-stu-id="a5161-117">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="a5161-118">Изменить значение другой ячейки.</span><span class="sxs-lookup"><span data-stu-id="a5161-118">Change another cell's value.</span></span>
- <span data-ttu-id="a5161-119">Перемещение, переименование, удаление и добавление листов в книгу.</span><span class="sxs-lookup"><span data-stu-id="a5161-119">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="a5161-120">Измените любые параметры среды, такие как режим вычисления или экранные представления.</span><span class="sxs-lookup"><span data-stu-id="a5161-120">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="a5161-121">Добавление имен в книгу.</span><span class="sxs-lookup"><span data-stu-id="a5161-121">Add names to a workbook.</span></span>
- <span data-ttu-id="a5161-122">Задайте свойства или выполните большинство методов.</span><span class="sxs-lookup"><span data-stu-id="a5161-122">Set properties or execute most methods.</span></span>

<span data-ttu-id="a5161-123">Изменение Excel может привести к ухудшению производительности, времени ожидания и бесконечному циклу.</span><span class="sxs-lookup"><span data-stu-id="a5161-123">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="a5161-124">Пользовательские вычисления функций не должны выполняться во время пересчета Excel, так как это приведет к непредсказуемым результатам.</span><span class="sxs-lookup"><span data-stu-id="a5161-124">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="a5161-125">Вместо этого внесите изменения в Excel из контекста кнопки на ленте или области задач.</span><span class="sxs-lookup"><span data-stu-id="a5161-125">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a5161-126">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="a5161-126">Next steps</span></span>

- [<span data-ttu-id="a5161-127">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a5161-127">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="a5161-128">См. также</span><span class="sxs-lookup"><span data-stu-id="a5161-128">See also</span></span>

- [<span data-ttu-id="a5161-129">Обмен данными и событиями между пользовательскими функциями и областью задач Excel</span><span class="sxs-lookup"><span data-stu-id="a5161-129">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
