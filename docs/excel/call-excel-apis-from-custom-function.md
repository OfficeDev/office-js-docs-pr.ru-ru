---
title: Вызов API JavaScript Excel из настраиваемой функции
description: Узнайте, какие API JavaScript Excel можно вызвать из настраиваемой функции.
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613908"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a><span data-ttu-id="00486-103">Вызов API JavaScript Excel из настраиваемой функции</span><span class="sxs-lookup"><span data-stu-id="00486-103">Call Excel JavaScript APIs from a custom function</span></span>

<span data-ttu-id="00486-104">Вызов API JavaScript Excel из пользовательских функций, чтобы получить данные о диапазоне и получить дополнительный контекст для вычислений.</span><span class="sxs-lookup"><span data-stu-id="00486-104">Call Excel JavaScript APIs from your custom functions to get range data and obtain more context for your calculations.</span></span> <span data-ttu-id="00486-105">Вызов API JavaScript Excel с помощью настраиваемой функции может быть полезен, если:</span><span class="sxs-lookup"><span data-stu-id="00486-105">Calling Excel JavaScript APIs through a custom function can be helpful when:</span></span>

- <span data-ttu-id="00486-106">Перед вычислением настраиваемая функция должна получать сведения из Excel.</span><span class="sxs-lookup"><span data-stu-id="00486-106">A custom function needs to get information from Excel before calculation.</span></span> <span data-ttu-id="00486-107">Эти сведения могут включать свойства документов, форматы диапазона, пользовательские XML-части, имя книги или другую информацию, определенную в Excel.</span><span class="sxs-lookup"><span data-stu-id="00486-107">This information might include document properties, range formats, custom XML parts, a workbook name, or other Excel-specific information.</span></span>
- <span data-ttu-id="00486-108">Настраиваемая функция будет устанавливать формат номера ячейки для возвращаемого значения после вычисления.</span><span class="sxs-lookup"><span data-stu-id="00486-108">A custom function will set the cell's number format for the return values after calculation.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="00486-109">Чтобы вызвать API JavaScript Excel из настраиваемой функции, необходимо использовать общее время запуска JavaScript.</span><span class="sxs-lookup"><span data-stu-id="00486-109">To call Excel JavaScript APIs from your custom function, you'll need to use a shared JavaScript runtime.</span></span> <span data-ttu-id="00486-110">Дополнительные сведения см. в статье [Настройка надстройки Office для использования общей среды выполнения JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="00486-110">See [Configure your Office Add-in to use a shared JavaScript runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) to learn more.</span></span>

## <a name="code-sample"></a><span data-ttu-id="00486-111">Пример кода</span><span class="sxs-lookup"><span data-stu-id="00486-111">Code sample</span></span>

<span data-ttu-id="00486-112">Чтобы вызвать API JavaScript Excel из настраиваемой функции, сначала требуется контекст.</span><span class="sxs-lookup"><span data-stu-id="00486-112">To call Excel JavaScript APIs from a custom function, you first need a context.</span></span> <span data-ttu-id="00486-113">Чтобы получить контекст, используйте объект [Excel.RequestContext.](/javascript/api/excel/excel.requestcontext)</span><span class="sxs-lookup"><span data-stu-id="00486-113">Use the [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) object to get a context.</span></span> <span data-ttu-id="00486-114">Затем используйте контекст для вызова API, необходимых в книге.</span><span class="sxs-lookup"><span data-stu-id="00486-114">Then use the context to call the APIs you need in the workbook.</span></span>

<span data-ttu-id="00486-115">В следующем примере кода показано, как использовать для получения значения из `Excel.RequestContext` ячейки в книге.</span><span class="sxs-lookup"><span data-stu-id="00486-115">The following code sample shows how to use `Excel.RequestContext` to get a value from a cell in the workbook.</span></span> <span data-ttu-id="00486-116">В этом примере параметр передается в метод `address` Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) и должен быть введен в качестве строки.</span><span class="sxs-lookup"><span data-stu-id="00486-116">In this sample, the `address` parameter is passed into the Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) method and must be entered as a string.</span></span> <span data-ttu-id="00486-117">Например, настраиваемая функция, вступив в пользовательский интерфейс Excel, должна следовать шаблону , где находится адрес ячейки, из которой можно получить `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` значение.</span><span class="sxs-lookup"><span data-stu-id="00486-117">For example, the custom function entered into the Excel UI must follow the pattern `=CONTOSO.GETRANGEVALUE("A1")`, where `"A1"` is the address of the cell from which to retrieve the value.</span></span>

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a><span data-ttu-id="00486-118">Ограничения вызовов API JavaScript Excel с помощью настраиваемой функции</span><span class="sxs-lookup"><span data-stu-id="00486-118">Limitations of calling Excel JavaScript APIs through a custom function</span></span>

<span data-ttu-id="00486-119">Не вызывайте API JavaScript Excel из настраиваемой функции, которая меняет среду Excel.</span><span class="sxs-lookup"><span data-stu-id="00486-119">Don't call Excel JavaScript APIs from a custom function that change the environment of Excel.</span></span> <span data-ttu-id="00486-120">Это означает, что пользовательские функции не должны выполнять следующие функции:</span><span class="sxs-lookup"><span data-stu-id="00486-120">This means your custom functions should not do any of the following:</span></span>

- <span data-ttu-id="00486-121">Вставка, удаление или форматирование ячеек в таблицу.</span><span class="sxs-lookup"><span data-stu-id="00486-121">Insert, delete, or format cells on the spreadsheet.</span></span>
- <span data-ttu-id="00486-122">Измените значение другой ячейки.</span><span class="sxs-lookup"><span data-stu-id="00486-122">Change another cell's value.</span></span>
- <span data-ttu-id="00486-123">Перемещение, переименование, удаление или добавление листов в книгу.</span><span class="sxs-lookup"><span data-stu-id="00486-123">Move, rename, delete, or add sheets to a workbook.</span></span>
- <span data-ttu-id="00486-124">Измените все параметры среды, такие как режим вычисления или представления экрана.</span><span class="sxs-lookup"><span data-stu-id="00486-124">Change any of the environment options, such as calculation mode or screen views.</span></span>
- <span data-ttu-id="00486-125">Добавление имен в книгу.</span><span class="sxs-lookup"><span data-stu-id="00486-125">Add names to a workbook.</span></span>
- <span data-ttu-id="00486-126">Установите свойства или выполните большинство методов.</span><span class="sxs-lookup"><span data-stu-id="00486-126">Set properties or execute most methods.</span></span>

<span data-ttu-id="00486-127">Изменение Excel может привести к низкой производительности, выходу времени и бесконечным циклам.</span><span class="sxs-lookup"><span data-stu-id="00486-127">Changing Excel can result in poor performance, time outs, and infinite loops.</span></span> <span data-ttu-id="00486-128">Настраиваемые вычисления функций не должны запускаться во время пересчета Excel, так как это приведет к непредсказуемым результатам.</span><span class="sxs-lookup"><span data-stu-id="00486-128">Custom function calculations shouldn't run while an Excel recalculation is taking place as it will result in unpredictable results.</span></span>

<span data-ttu-id="00486-129">Вместо этого внести изменения в Excel из контекста кнопки ленты или области задач.</span><span class="sxs-lookup"><span data-stu-id="00486-129">Instead, make changes to Excel from the context of a ribbon button, or task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="00486-130">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="00486-130">Next steps</span></span>

- [<span data-ttu-id="00486-131">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="00486-131">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a><span data-ttu-id="00486-132">См. также</span><span class="sxs-lookup"><span data-stu-id="00486-132">See also</span></span>

- [<span data-ttu-id="00486-133">Совместное делиться данными и событиями между пользовательскими функциями Excel и учебником по области задач</span><span class="sxs-lookup"><span data-stu-id="00486-133">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="00486-134">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="00486-134">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
