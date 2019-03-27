---
title: Работа с листами с использованием API JavaScript для Excel
description: ''
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: 825ae88afd98afbcd268716c93ddcb13d24a9a1e
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871558"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="39129-102">Работа с листами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="39129-102">Work with worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="39129-p101">В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для листов с использованием API JavaScript для Excel. Полный список свойств и методов, поддерживаемых объектами **Worksheet** и **WorksheetCollection**, см. в статьях [Объект Worksheet (API JavaScript для Excel)](/javascript/api/excel/excel.worksheet) и [Объект WorksheetCollection (API JavaScript для Excel)](/javascript/api/excel/excel.worksheetcollection).</span><span class="sxs-lookup"><span data-stu-id="39129-p101">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API. For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="39129-105">Сведения в этой статье применимы только к обычным листам, а не к листам диаграмм или макросов.</span><span class="sxs-lookup"><span data-stu-id="39129-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="39129-106">Получение листов</span><span class="sxs-lookup"><span data-stu-id="39129-106">Get worksheets</span></span>

<span data-ttu-id="39129-107">В примере ниже показано, как возвратить коллекцию листов, загрузить свойство **name** каждого листа и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="39129-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length > 1) {
                console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
            } else {
                console.log(`There is one worksheet in the workbook:`);
            }
            for (var i in sheets.items) {
                console.log(sheets.items[i].name);
            }
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="39129-p102">Свойство **id** листа уникальным образом идентифицирует лист в конкретной книге, и его значение не изменяется даже при переименовании или перемещении листа. При удалении листа из книги в Excel для Mac **идентификатор** удаленного листа можно назначить новому листу (созданному после удаления).</span><span class="sxs-lookup"><span data-stu-id="39129-p102">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved. When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="39129-110">Получение активного листа</span><span class="sxs-lookup"><span data-stu-id="39129-110">Get the active worksheet</span></span>

<span data-ttu-id="39129-111">В примере кода ниже показано, как получить активный лист, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="39129-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-the-active-worksheet"></a><span data-ttu-id="39129-112">Задание активного листа</span><span class="sxs-lookup"><span data-stu-id="39129-112">Set the active worksheet</span></span>

<span data-ttu-id="39129-p103">В примере кода ниже показано, как задать лист **Sample** (Пример) в качестве активного, загрузить его свойство **name** и записать сообщение в консоль. Если нет листа с таким именем, метод **activate()** создаст ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="39129-p103">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console. If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.activate();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="39129-115">Ссылка на листы по их относительным положениям</span><span class="sxs-lookup"><span data-stu-id="39129-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="39129-116">В примерах ниже показано, как ссылаться на лист по его относительному положению.</span><span class="sxs-lookup"><span data-stu-id="39129-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="39129-117">Получение первого листа</span><span class="sxs-lookup"><span data-stu-id="39129-117">Get the first worksheet</span></span>

<span data-ttu-id="39129-118">В примере кода ниже показано, как получить первый лист в книге, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="39129-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var firstSheet = context.workbook.worksheets.getFirst();
    firstSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the first worksheet is "${firstSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-last-worksheet"></a><span data-ttu-id="39129-119">Получение последнего листа</span><span class="sxs-lookup"><span data-stu-id="39129-119">Get the last worksheet</span></span>

<span data-ttu-id="39129-120">В примере кода ниже показано, как получить последний лист в книге, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="39129-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var lastSheet = context.workbook.worksheets.getLast();
    lastSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the last worksheet is "${lastSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-next-worksheet"></a><span data-ttu-id="39129-121">Получение следующего листа</span><span class="sxs-lookup"><span data-stu-id="39129-121">Get the next worksheet</span></span>

<span data-ttu-id="39129-p104">В примере кода ниже показано, как получить лист, следующий за активным листом, в книге, загрузить его свойство **name** и записать сообщение в консоль. Если нет листа после активного листа, метод **getNext()** создаст ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="39129-p104">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

```js
 Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var nextSheet = currentSheet.getNext();
    nextSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that follows the active worksheet is "${nextSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="39129-124">Получение предыдущего листа</span><span class="sxs-lookup"><span data-stu-id="39129-124">Get the previous worksheet</span></span>

<span data-ttu-id="39129-p105">В примере кода ниже показано, как получить лист, предшествующий активному листу, в книге, загрузить его свойство **name** и записать сообщение в консоль. Если нет листа перед активным листом, метод **getPrevious()** создаст ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="39129-p105">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var previousSheet = currentSheet.getPrevious();
    previousSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that precedes the active worksheet is "${previousSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="add-a-worksheet"></a><span data-ttu-id="39129-127">Добавление листа</span><span class="sxs-lookup"><span data-stu-id="39129-127">Add a worksheet</span></span>

<span data-ttu-id="39129-p106">В примере кода ниже показано, как добавить лист **Sample** (Пример) в рабочую книгу, загрузить его свойства **name** и **position** и записать сообщение в консоль. Новый лист будет следовать за всеми остальными.</span><span class="sxs-lookup"><span data-stu-id="39129-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;

    var sheet = sheets.add("Sample");
    sheet.load("name, position");

    return context.sync()
        .then(function () {
            console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
        });
}).catch(errorHandlerFunction);
```

## <a name="delete-a-worksheet"></a><span data-ttu-id="39129-130">Удаление листа</span><span class="sxs-lookup"><span data-stu-id="39129-130">Delete a worksheet</span></span>

<span data-ttu-id="39129-131">В примере кода ниже показано, как удалить последний лист в книге (если это не единственный лист в книге) и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="39129-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length === 1) {
                console.log("Unable to delete the only worksheet in the workbook");
            } else {
                var lastSheet = sheets.items[sheets.items.length - 1];

                console.log(`Deleting worksheet named "${lastSheet.name}"`);
                lastSheet.delete();

                return context.sync();
            };
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="39129-132">Лист с уровнем скрытия "[надежно скрыт](/javascript/api/excel/excel.sheetvisibility)" невозможно удалить с помощью метода `delete`.</span><span class="sxs-lookup"><span data-stu-id="39129-132">A worksheet with a visibility of "[Very Hidden](/javascript/api/excel/excel.sheetvisibility)" cannot be deleted with the `delete` method.</span></span> <span data-ttu-id="39129-133">Чтобы удалить лист, нужно сперва изменить его уровень скрытия.</span><span class="sxs-lookup"><span data-stu-id="39129-133">If you wish to delete the worksheet anyway, you must first change the visibility.</span></span>

## <a name="rename-a-worksheet"></a><span data-ttu-id="39129-134">Переименование листа</span><span class="sxs-lookup"><span data-stu-id="39129-134">Rename a worksheet</span></span>

<span data-ttu-id="39129-135">В примере ниже показано, как изменить имя активного листа на **New Name** (Новое имя).</span><span class="sxs-lookup"><span data-stu-id="39129-135">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="39129-136">Перемещение листа</span><span class="sxs-lookup"><span data-stu-id="39129-136">Move a worksheet</span></span>

<span data-ttu-id="39129-137">В примере ниже показано, как переместить лист из последней позиции в книге на первую.</span><span class="sxs-lookup"><span data-stu-id="39129-137">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items");

    return context.sync()
        .then(function () {
            var lastSheet = sheets.items[sheets.items.length - 1];
            lastSheet.position = 0;

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

## <a name="set-worksheet-visibility"></a><span data-ttu-id="39129-138">Настройка видимости листа</span><span class="sxs-lookup"><span data-stu-id="39129-138">Set worksheet visibility</span></span>

<span data-ttu-id="39129-139">В примерах ниже показано, как настроить видимость листа.</span><span class="sxs-lookup"><span data-stu-id="39129-139">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="39129-140">Скрытие листа</span><span class="sxs-lookup"><span data-stu-id="39129-140">Hide a worksheet</span></span>

<span data-ttu-id="39129-141">В примере кода ниже показано, как сделать лист **Sample** (Пример) скрытым, загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="39129-141">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.hidden;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is hidden`);
        });
}).catch(errorHandlerFunction);
```

### <a name="unhide-a-worksheet"></a><span data-ttu-id="39129-142">Отмена скрытия листа</span><span class="sxs-lookup"><span data-stu-id="39129-142">Unhide a worksheet</span></span>

<span data-ttu-id="39129-143">В примере кода ниже показано, как сделать лист **Sample** (Пример), загрузить его свойство **name** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="39129-143">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is visible`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-a-single-cell-within-a-worksheet"></a><span data-ttu-id="39129-144">Получение одной ячейки листа</span><span class="sxs-lookup"><span data-stu-id="39129-144">Get a single cell within a worksheet</span></span>

<span data-ttu-id="39129-145">В примере кода ниже показано, как получить ячейку, расположенную в строке 2 и столбце 5 листа **Sample** (Пример), загрузить его свойства **address** и **values** и записать сообщение в консоль.</span><span class="sxs-lookup"><span data-stu-id="39129-145">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="39129-146">Значения, передаваемые в метод `getCell(row: number, column:number)`, представляют собой индексируемые с нуля номера строк и столбцов получаемой ячейки.</span><span class="sxs-lookup"><span data-stu-id="39129-146">The values that are passed into the `getCell(row: number, column:number)` method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var cell = sheet.getCell(1, 4);
    cell.load("address, values");

    return context.sync()
        .then(function() {
            console.log(`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`);
        })
}).catch(errorHandlerFunction);
```

## <a name="find-all-cells-with-matching-text-preview"></a><span data-ttu-id="39129-147">Поиск всех ячеек с соответствующим текстом (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="39129-147">Find all cells with matching text (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="39129-148">Функция `findAll` объекта Worksheet в настоящее время доступна только в общедоступной предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="39129-148">The Worksheet object's `findAll` function is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="39129-149">У объекта `Worksheet` есть метод `find` для поиска указанной строки в листе.</span><span class="sxs-lookup"><span data-stu-id="39129-149">The `Worksheet` object has a `find` method to search for a specified string within the worksheet.</span></span> <span data-ttu-id="39129-150">Он возвращает объект `RangeAreas`, являющийся коллекцией объектов `Range`, которые можно отредактировать все сразу.</span><span class="sxs-lookup"><span data-stu-id="39129-150">It returns a `RangeAreas` object, which is a collection of `Range` objects that can be edited all at once.</span></span> <span data-ttu-id="39129-151">Приведенный ниже пример кода находит все ячейки со значениями, соответствующими строке **Complete** (Завершено), и окрашивает их зеленым цветом.</span><span class="sxs-lookup"><span data-stu-id="39129-151">The following code sample finds all cells with values equal to the string **Complete** and colors them green.</span></span> <span data-ttu-id="39129-152">Обратите внимание, что метод `findAll` выдаст ошибку `ItemNotFound`, если указанной строки не существует в листе.</span><span class="sxs-lookup"><span data-stu-id="39129-152">Note that `findAll` will throw an `ItemNotFound` error if the specified string doesn't exist in the worksheet.</span></span> <span data-ttu-id="39129-153">Если ожидается, что указанная строка может отсутствовать в листе, используйте вместо этого метод [findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods), чтобы ваш код корректно обработал этот сценарий.</span><span class="sxs-lookup"><span data-stu-id="39129-153">If you expect that the specified string may not exist in the worksheet, use the [findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var foundRanges = sheet.findAll("Complete", {
        completeMatch: true, // findAll will match the whole cell value
        matchCase: false // findAll will not match case
    });

    return context.sync()
        .then(function() {
            foundRanges.format.fill.color = "green"
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="39129-154">В этом разделе описано, как найти ячейки и диапазоны с помощью функций объекта `Worksheet`.</span><span class="sxs-lookup"><span data-stu-id="39129-154">This section describes how to find cells and ranges using the `Worksheet` object's functions.</span></span> <span data-ttu-id="39129-155">Дополнительные сведения об извлечении диапазонов можно найти в статьях о конкретных объектах.</span><span class="sxs-lookup"><span data-stu-id="39129-155">More range retrieval information can be found in object-specific articles.</span></span>
> - <span data-ttu-id="39129-156">Примеры, в которых показано, как получить диапазон в листе с помощью объекта `Range`, см. в статье [Работа с диапазонами с использованием API JavaScript для Excel](excel-add-ins-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="39129-156">For examples that show how to get a range within a worksheet using the `Range` object, see [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>
> - <span data-ttu-id="39129-157">Примеры, в которых показано, как получить диапазоны из объекта `Table`, см. в статье [Работа с таблицами с использованием API JavaScript для Excel](excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="39129-157">For examples that show how to get ranges from a `Table` object, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>
> - <span data-ttu-id="39129-158">Примеры, в которых показано, как выполнять поиск большого диапазона для нескольких поддиапазонов с учетом характеристик ячеек, см. в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="39129-158">For examples that show how to search a large range for multiple sub-ranges based on cell characteristics, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="data-protection"></a><span data-ttu-id="39129-159">Защита данных</span><span class="sxs-lookup"><span data-stu-id="39129-159">Data protection</span></span>

<span data-ttu-id="39129-160">Надстройка может управлять возможностью пользователя по изменению данных на листе.</span><span class="sxs-lookup"><span data-stu-id="39129-160">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="39129-161">Свойство `protection` листа является объектом [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) с методом `protect()`.</span><span class="sxs-lookup"><span data-stu-id="39129-161">The worksheet's `protection` property is a [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="39129-162">В приведенном ниже примере показан основной сценарий переключения полной защиты активного листа.</span><span class="sxs-lookup"><span data-stu-id="39129-162">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("protection/protected");

    return context.sync().then(function() {
        if (!activeSheet.protection.protected) {
            activeSheet.protection.protect();
        }
    })
}).catch(errorHandlerFunction);
```

<span data-ttu-id="39129-163">Метод `protect` содержит два необязательных параметра:</span><span class="sxs-lookup"><span data-stu-id="39129-163">The `protect` method has two optional parameters:</span></span>

- <span data-ttu-id="39129-164">`options`: объект [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions), определяющий конкретные ограничения на редактирование.</span><span class="sxs-lookup"><span data-stu-id="39129-164">`options`: A [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
- <span data-ttu-id="39129-165">`password`: строка, представляющая пароль, необходимый пользователю для обхода защиты и редактирования листа.</span><span class="sxs-lookup"><span data-stu-id="39129-165">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="39129-166">В статье [Защита листа](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) содержатся дополнительные сведения о защите листа и ее изменении с помощью пользовательского интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="39129-166">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="see-also"></a><span data-ttu-id="39129-167">См. также</span><span class="sxs-lookup"><span data-stu-id="39129-167">See also</span></span>

- [<span data-ttu-id="39129-168">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="39129-168">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
