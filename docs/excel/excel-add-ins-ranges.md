---
title: Работа с диапазонами с использованием API JavaScript для Excel (основные задачи)
description: ''
ms.date: 12/28/2018
ms.openlocfilehash: 843f57f8e5dc20d4341749f4594e0bd8139e60fa
ms.sourcegitcommit: d75295cc4f47d8d872e7a361fdb5526f0f145dd2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/29/2018
ms.locfileid: "27460872"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="1b876-102">Работа с диапазонами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="1b876-102">Work with ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="1b876-103">В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для диапазонов с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="1b876-103">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="1b876-104">Полный список свойств и методов, поддерживаемых объектом **Range**, см. в статье [Объект Range (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="1b876-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

> [!NOTE]
> <span data-ttu-id="1b876-105">Примеры кода, в которых показано, как выполнять более сложные задачи для диапазонов, см. в статье [Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)](excel-add-ins-ranges-advanced.md).</span><span class="sxs-lookup"><span data-stu-id="1b876-105">For code samples that show how to perform more advanced tasks with ranges, see [Work with ranges using the Excel JavaScript API (advanced)](excel-add-ins-ranges-advanced.md).</span></span>

## <a name="get-a-range"></a><span data-ttu-id="1b876-106">Получение диапазона</span><span class="sxs-lookup"><span data-stu-id="1b876-106">Get a range</span></span>

<span data-ttu-id="1b876-107">В примерах ниже показаны различные способы получения ссылки на диапазон, расположенный на листе.</span><span class="sxs-lookup"><span data-stu-id="1b876-107">The following examples show different ways to get a reference to a range within a worksheet.</span></span>

### <a name="get-range-by-address"></a><span data-ttu-id="1b876-108">Получение диапазона по адресу</span><span class="sxs-lookup"><span data-stu-id="1b876-108">Get range by address</span></span>

<span data-ttu-id="1b876-109">В примере кода ниже показано, как получить диапазон с адресом **B2:B5** с листа **Sample** (Пример), загрузить его свойство **address** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="1b876-109">The following code sample gets the range with address **B2:B5** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-range-by-name"></a><span data-ttu-id="1b876-110">Получение диапазона по имени</span><span class="sxs-lookup"><span data-stu-id="1b876-110">Get range by name</span></span>

<span data-ttu-id="1b876-111">В примере кода ниже показано, как получить диапазон с именем **MyRange** (Мой диапазон) с листа **Sample** (Пример), загрузить его свойство **address** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="1b876-111">The following code sample gets the range named **MyRange** from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-used-range"></a><span data-ttu-id="1b876-112">Получение используемого диапазона</span><span class="sxs-lookup"><span data-stu-id="1b876-112">Get used range</span></span>

<span data-ttu-id="1b876-113">В примере кода ниже показано, как получить используемый диапазон с листа **Sample** (Пример), загрузить его свойство **address** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="1b876-113">The following code sample gets the used range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span> <span data-ttu-id="1b876-114">Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки листа, которые содержат значение или форматирование.</span><span class="sxs-lookup"><span data-stu-id="1b876-114">The used range is the smallest range that encompasses any cells in the worksheet that have a value or formatting assigned to them.</span></span> <span data-ttu-id="1b876-115">Если весь лист пуст, метод **getUsedRange()** возвращает диапазон, состоящий только из левой верхней ячейки листа.</span><span class="sxs-lookup"><span data-stu-id="1b876-115">If the entire worksheet is blank, the **getUsedRange()** method returns a range that consists of only the top-left cell in the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-entire-range"></a><span data-ttu-id="1b876-116">Получение всего диапазона</span><span class="sxs-lookup"><span data-stu-id="1b876-116">Get entire range</span></span>

<span data-ttu-id="1b876-117">В примере кода ниже показано, как получить весь диапазон листа **Sample** (Пример), загрузить его свойство **address** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="1b876-117">The following code sample gets the entire worksheet range from the worksheet named **Sample**, loads its **address** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="insert-a-range-of-cells"></a><span data-ttu-id="1b876-118">Вставка диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="1b876-118">Insert a range of cells</span></span>

<span data-ttu-id="1b876-119">В примере кода ниже показано, как вставить диапазон ячеек в расположение **B4:E4** и сдвинуть другие ячейки вниз, чтобы освободить место для новых ячеек.</span><span class="sxs-lookup"><span data-stu-id="1b876-119">The following code sample inserts a range of cells in location **B4:E4** and shifts other cells down to provide space for the new cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-120">**Данные перед вставкой диапазона**</span><span class="sxs-lookup"><span data-stu-id="1b876-120">**Data before range is inserted**</span></span>

![Данные в Excel перед вставкой диапазона](../images/excel-ranges-start.png)

<span data-ttu-id="1b876-122">**Данные после вставки диапазона**</span><span class="sxs-lookup"><span data-stu-id="1b876-122">**Data after range is inserted**</span></span>

![Данные в Excel после вставки диапазона](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a><span data-ttu-id="1b876-124">Очистка диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="1b876-124">Clear a range of cells</span></span>

<span data-ttu-id="1b876-125">В примере кода ниже показано, как удалить все содержимое и форматирование ячеек в диапазоне **E2:E5**.</span><span class="sxs-lookup"><span data-stu-id="1b876-125">The following code sample clears all contents and formatting of cells in the range **E2:E5**.</span></span>  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-126">**Данные перед очисткой диапазона**</span><span class="sxs-lookup"><span data-stu-id="1b876-126">**Data before range is cleared**</span></span>

![Данные в Excel перед очисткой диапазона](../images/excel-ranges-start.png)

<span data-ttu-id="1b876-128">**Данные после очистки диапазона**</span><span class="sxs-lookup"><span data-stu-id="1b876-128">**Data after range is cleared**</span></span>

![Данные в Excel после очистки диапазона](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a><span data-ttu-id="1b876-130">Удаление диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="1b876-130">Delete a range of cells</span></span>

<span data-ttu-id="1b876-131">В примере кода ниже показано, как удалить ячейки в диапазоне **B4:E4** и сдвинуть другие ячейки вверх, чтобы заполнить место, освободившееся после удаления ячеек.</span><span class="sxs-lookup"><span data-stu-id="1b876-131">The following code sample deletes the cells in the range **B4:E4** and shift other cells up to fill the space that was vacated by the deleted cells.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-132">**Данные перед удалением диапазона**</span><span class="sxs-lookup"><span data-stu-id="1b876-132">**Data before range is deleted**</span></span>

![Данные в Excel перед удалением диапазона](../images/excel-ranges-start.png)

<span data-ttu-id="1b876-134">**Данные после удаления диапазона**</span><span class="sxs-lookup"><span data-stu-id="1b876-134">**Data after range is deleted**</span></span>

![Данные в Excel после удаления диапазона](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a><span data-ttu-id="1b876-136">Задание выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="1b876-136">Set the selected range</span></span>

<span data-ttu-id="1b876-137">В примере кода ниже показано, как выделить диапазон **B2:E6** на активном листе.</span><span class="sxs-lookup"><span data-stu-id="1b876-137">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-138">**Выделенный диапазон B2:E6**</span><span class="sxs-lookup"><span data-stu-id="1b876-138">**Selected range B2:E6**</span></span>

![Выделенный диапазон в Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="1b876-140">Получение выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="1b876-140">Get the selected range</span></span>

<span data-ttu-id="1b876-141">В примере кода ниже показано, как получить выделенный диапазон, загрузить его свойство **address** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="1b876-141">The following code sample gets the selected range, loads its **address** property, and writes a message to the console.</span></span> 

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-values-or-formulas"></a><span data-ttu-id="1b876-142">Задание значений или формул</span><span class="sxs-lookup"><span data-stu-id="1b876-142">Set values or formulas</span></span>

<span data-ttu-id="1b876-143">В примерах ниже показано, как задать значения и формулы для одной ячейки или диапазона ячеек.</span><span class="sxs-lookup"><span data-stu-id="1b876-143">The following examples show how to set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="1b876-144">Задание значения для одной ячейки</span><span class="sxs-lookup"><span data-stu-id="1b876-144">Set value for a single cell</span></span>

<span data-ttu-id="1b876-145">В примере кода ниже показано, как присвоить ячейке **C3** значение 5, а затем настроить ширину столбцов для наилучшего размещения данных.</span><span class="sxs-lookup"><span data-stu-id="1b876-145">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-146">**Данные перед изменением значения ячейки**</span><span class="sxs-lookup"><span data-stu-id="1b876-146">**Data before cell value is updated**</span></span>

![Данные в Excel перед изменением значения ячейки](../images/excel-ranges-set-start.png)

<span data-ttu-id="1b876-148">**Данные после изменения значения ячейки**</span><span class="sxs-lookup"><span data-stu-id="1b876-148">**Data after cell value is updated**</span></span>

![Данные в Excel после изменения значения ячейки](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="1b876-150">Задание значений для диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="1b876-150">Set values for a range of cells</span></span>

<span data-ttu-id="1b876-151">В примере кода ниже показано, как присвоить значения ячейкам в диапазоне **B5:D5**, а затем настроить ширину столбцов для наилучшего размещения данных.</span><span class="sxs-lookup"><span data-stu-id="1b876-151">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];
    
    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-152">**Данные перед изменением значений ячеек**</span><span class="sxs-lookup"><span data-stu-id="1b876-152">**Data before cell values are updated**</span></span>

![Данные в Excel перед изменением значений ячеек](../images/excel-ranges-set-start.png)

<span data-ttu-id="1b876-154">**Данные после изменения значений ячеек**</span><span class="sxs-lookup"><span data-stu-id="1b876-154">**Data after cell values are updated**</span></span>

![Данные в Excel после изменения значений ячеек](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="1b876-156">Задание формулы для одной ячейки</span><span class="sxs-lookup"><span data-stu-id="1b876-156">Set formula for a single cell</span></span>

<span data-ttu-id="1b876-157">В примере кода ниже показано, как задать формулу для ячейки **E3**, а затем настроить ширину столбцов для наилучшего размещения данных.</span><span class="sxs-lookup"><span data-stu-id="1b876-157">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-158">**Данные перед заданием формулы для ячейки**</span><span class="sxs-lookup"><span data-stu-id="1b876-158">**Data before cell formula is set**</span></span>

![Данные в Excel перед заданием формулы для ячейки](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="1b876-160">**Данные после задания формулы для ячейки**</span><span class="sxs-lookup"><span data-stu-id="1b876-160">**Data after cell formula is set**</span></span>

![Данные в Excel после задания формулы для ячейки](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="1b876-162">Задание формул для диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="1b876-162">Set formulas for a range of cells</span></span>

<span data-ttu-id="1b876-163">В примере кода ниже показано, как задать формулы для ячеек в диапазоне **E2:E6**, а затем настроить ширину столбцов для наилучшего размещения данных.</span><span class="sxs-lookup"><span data-stu-id="1b876-163">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];
    
    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-164">**Данные перед заданием формул для ячеек**</span><span class="sxs-lookup"><span data-stu-id="1b876-164">**Data before cell formulas are set**</span></span>

![Данные в Excel перед заданием формул для ячеек](../images/excel-ranges-start-set-formula.png)

<span data-ttu-id="1b876-166">**Данные после задания формул для ячеек**</span><span class="sxs-lookup"><span data-stu-id="1b876-166">**Data after cell formulas are set**</span></span>

![Данные в Excel после задания формул для ячеек](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="1b876-168">Получение значений, текста или формул</span><span class="sxs-lookup"><span data-stu-id="1b876-168">Get values, text, or formulas</span></span>

<span data-ttu-id="1b876-169">В примерах ниже показано, как получать значения, текст и формулы из диапазона ячеек.</span><span class="sxs-lookup"><span data-stu-id="1b876-169">These examples show how to get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="1b876-170">Получение значений из диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="1b876-170">Get values from a range of cells</span></span>

<span data-ttu-id="1b876-171">В примере кода ниже показано, как получить диапазон **B2:E6**, загрузить его свойство **values** и записать значения из этого свойства в консоль.</span><span class="sxs-lookup"><span data-stu-id="1b876-171">The following code sample gets the range **B2:E6**, loads its **values** property, and writes the values to the console.</span></span> <span data-ttu-id="1b876-172">Свойство **values** диапазона указывает необработанные значения, содержащиеся в ячейках.</span><span class="sxs-lookup"><span data-stu-id="1b876-172">The **values** property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="1b876-173">Даже если некоторые ячейки в диапазоне содержат формулы, свойство **values** диапазона будет указывать необработанные значения для этих ячеек, а не формулы.</span><span class="sxs-lookup"><span data-stu-id="1b876-173">Even if some cells in a range contain formulas, the **values** property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-174">**Данные в диапазоне (значения в столбце E представляют собой результат вычисления формул)**</span><span class="sxs-lookup"><span data-stu-id="1b876-174">**Data in range (values in column E are a result of formulas)**</span></span>

![Данные в Excel после задания формул для ячеек](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="1b876-176">**range.values (как записано в консоль в примере кода выше)**</span><span class="sxs-lookup"><span data-stu-id="1b876-176">**range.values (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="1b876-177">Получение текста из диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="1b876-177">Get text from a range of cells</span></span>

<span data-ttu-id="1b876-178">В примере кода ниже показано, как получить диапазон **B2:E6**, загрузить его свойство **text** и записать текст из этого свойства в консоль.</span><span class="sxs-lookup"><span data-stu-id="1b876-178">The following code sample gets the range **B2:E6**, loads its **text** property, and writes it to the console.</span></span>  <span data-ttu-id="1b876-179">Свойство **text** диапазона указывает отображаемые значения для ячеек в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="1b876-179">The **text** property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="1b876-180">Даже если некоторые ячейки в диапазоне содержат формулы, свойство **text** диапазона будет указывать отображаемые значения для этих ячеек, а не формулы.</span><span class="sxs-lookup"><span data-stu-id="1b876-180">Even if some cells in a range contain formulas, the **text** property of the range specifies the display values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-181">**Данные в диапазоне (значения в столбце E представляют собой результат вычисления формул)**</span><span class="sxs-lookup"><span data-stu-id="1b876-181">**Data in range (values in column E are a result of formulas)**</span></span>

![Данные в Excel после задания формул для ячеек](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="1b876-183">**range.text (как записано в консоль в примере кода выше)**</span><span class="sxs-lookup"><span data-stu-id="1b876-183">**range.text (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="1b876-184">Получение формул из диапазона ячеек</span><span class="sxs-lookup"><span data-stu-id="1b876-184">Get formulas from a range of cells</span></span>

<span data-ttu-id="1b876-185">В примере кода ниже показано, как получить диапазон **B2:E6**, загрузить его свойство **formulas** и записать содержимое этого свойства в консоль.</span><span class="sxs-lookup"><span data-stu-id="1b876-185">The following code sample gets the range **B2:E6**, loads its **formulas** property, and writes it to the console.</span></span>  <span data-ttu-id="1b876-186">Свойство **formulas** диапазона указывает формулы для ячеек, содержащих формулы, и необработанные значения для ячеек, не содержащих формулы, в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="1b876-186">The **formulas** property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-187">**Данные в диапазоне (значения в столбце E представляют собой результат вычисления формул)**</span><span class="sxs-lookup"><span data-stu-id="1b876-187">**Data in range (values in column E are a result of formulas)**</span></span>

![Данные в Excel после задания формул для ячеек](../images/excel-ranges-set-formulas.png)

<span data-ttu-id="1b876-189">**range.formulas (как записано в консоль в примере кода выше)**</span><span class="sxs-lookup"><span data-stu-id="1b876-189">**range.formulas (as logged to the console by the code sample above)**</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="set-range-format"></a><span data-ttu-id="1b876-190">Задание формата диапазона</span><span class="sxs-lookup"><span data-stu-id="1b876-190">Set range format</span></span>

<span data-ttu-id="1b876-191">В примерах ниже показано, как задать цвет шрифта, цвет заливки и формат чисел для ячеек в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="1b876-191">The following examples show how to set font color, fill color, and number format for cells in a range.</span></span>

### <a name="set-font-color-and-fill-color"></a><span data-ttu-id="1b876-192">Задание цвета шрифта и цвета заливки</span><span class="sxs-lookup"><span data-stu-id="1b876-192">Set font color and fill color</span></span>

<span data-ttu-id="1b876-193">В примере ниже показано, как задать цвет шрифта и цвет заливки для ячеек в диапазоне **B2: E2**.</span><span class="sxs-lookup"><span data-stu-id="1b876-193">The following code sample sets the font color and fill color for cells in range **B2:E2**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-194">**Данные в диапазоне перед заданием цвета шрифта и цвета заливки**</span><span class="sxs-lookup"><span data-stu-id="1b876-194">**Data in range before font color and fill color are set**</span></span>

![Данные в Excel перед заданием формата](../images/excel-ranges-format-before.png)

<span data-ttu-id="1b876-196">**Данные в диапазоне после задания цвета шрифта и цвета заливки**</span><span class="sxs-lookup"><span data-stu-id="1b876-196">**Data in range after font color and fill color are set**</span></span>

![Данные в Excel после задания формата](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a><span data-ttu-id="1b876-198">Задание формата чисел</span><span class="sxs-lookup"><span data-stu-id="1b876-198">Set number format</span></span>

<span data-ttu-id="1b876-199">В примере ниже показано, как задать формат чисел для ячеек в диапазоне **D3:E5**.</span><span class="sxs-lookup"><span data-stu-id="1b876-199">The following code sample sets the number format for the cells in range **D3:E5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-200">**Данные в диапазоне перед заданием формата чисел**</span><span class="sxs-lookup"><span data-stu-id="1b876-200">**Data in range before number format is set**</span></span>

![Данные в Excel перед заданием формата](../images/excel-ranges-format-font-and-fill.png)

<span data-ttu-id="1b876-202">**Данные в диапазоне после задания формата чисел**</span><span class="sxs-lookup"><span data-stu-id="1b876-202">**Data in range after number format is set**</span></span>

![Данные в Excel после задания формата](../images/excel-ranges-format-numbers.png)

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="1b876-204">Условное форматирование диапазонов</span><span class="sxs-lookup"><span data-stu-id="1b876-204">Conditional formatting of ranges</span></span>

<span data-ttu-id="1b876-205">В диапазонах может применяться форматирование к отдельным ячейкам на основе условий.</span><span class="sxs-lookup"><span data-stu-id="1b876-205">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="1b876-206">Дополнительные сведения об этом см. в статье [Применение условного форматирования к диапазонам Excel](excel-add-ins-conditional-formatting.md).</span><span class="sxs-lookup"><span data-stu-id="1b876-206">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="find-a-cell-using-string-matching-preview"></a><span data-ttu-id="1b876-207">Поиск ячейки с помощью сопоставления строк (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="1b876-207">Find a cell using string matching (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="1b876-208">Функция `find` объекта Range в настоящее время доступна только в общедоступной предварительной версии (бета-версии).</span><span class="sxs-lookup"><span data-stu-id="1b876-208">The Range object's `find` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="1b876-209">Для применения этой функции необходимо использовать бета-версию библиотеки в CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="1b876-209">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="1b876-210">Если вы используете TypeScript или ваш редактор кода использует файлы определения типа TypeScript для IntelliSense, воспользуйтесь https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="1b876-210">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="1b876-211">У объекта `Range` есть метод `find` для поиска указанной строки в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="1b876-211">The `Range` object has a `find` method to search for a specified string within the range.</span></span> <span data-ttu-id="1b876-212">Он возвращает диапазон первой ячейки с текстом, соответствующим критериям.</span><span class="sxs-lookup"><span data-stu-id="1b876-212">It returns the range of the first cell with matching text.</span></span> <span data-ttu-id="1b876-213">Приведенный ниже пример кода находит первую ячейку со значением, соответствующим строке **Food** (Еда), и заносит ее адрес в консоль.</span><span class="sxs-lookup"><span data-stu-id="1b876-213">The following code sample finds the first cell with a value equal to the string **Food** and logs its address to the console.</span></span> <span data-ttu-id="1b876-214">Обратите внимание, что метод `find` выдает ошибку `ItemNotFound`, если указанной строки не существует в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="1b876-214">Note that `find` throws an `ItemNotFound` error if the specified string doesn't exist in the range.</span></span> <span data-ttu-id="1b876-215">Если ожидается, что указанная строка может отсутствовать в диапазоне, используйте вместо этого метод [findOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods), чтобы ваш код корректно обработал этот сценарий.</span><span class="sxs-lookup"><span data-stu-id="1b876-215">If you expect that the specified string may not exist in the range, use the [findOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="1b876-216">Если метод `find` вызывается для диапазона, представляющего одну ячейку, поиск выполняется во всем листе.</span><span class="sxs-lookup"><span data-stu-id="1b876-216">When the `find` method is called on a range representing a single cell, the entire worksheet is searched.</span></span> <span data-ttu-id="1b876-217">Поиск начинается в этой ячейке и продолжается в направлении, которое определяется параметром `SearchCriteria.searchDirection`, охватывающим концы листа при необходимости.</span><span class="sxs-lookup"><span data-stu-id="1b876-217">The search begins at that cell and goes in the direction specified by `SearchCriteria.searchDirection`, wrapping around the ends of the worksheet if needed.</span></span>

## <a name="see-also"></a><span data-ttu-id="1b876-218">См. также</span><span class="sxs-lookup"><span data-stu-id="1b876-218">See also</span></span>

- [<span data-ttu-id="1b876-219">Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)</span><span class="sxs-lookup"><span data-stu-id="1b876-219">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
- [<span data-ttu-id="1b876-220">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="1b876-220">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)