---
title: Работа с книгами с использованием API JavaScript для Excel
description: ''
ms.date: 11/27/2018
ms.openlocfilehash: 1cfde9bfdf306e35f47595f936679d9fa6e1814e
ms.sourcegitcommit: 026437bd3819f4e9cd4153ebe60c98ab04e18f4e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/30/2018
ms.locfileid: "27002344"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a><span data-ttu-id="78ad1-102">Работа с книгами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="78ad1-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="78ad1-103">В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для книг с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="78ad1-103">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="78ad1-104">Полный список свойств и методов, поддерживаемых объектом **Workbook**, см. в статье [Объект Workbook (API JavaScript для Excel)](/javascript/api/excel/excel.workbook).</span><span class="sxs-lookup"><span data-stu-id="78ad1-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.workbook).</span></span> <span data-ttu-id="78ad1-105">В этой статье также рассматриваются действия на уровне книги, выполняемые с помощью объекта [Application](/javascript/api/excel/excel.application).</span><span class="sxs-lookup"><span data-stu-id="78ad1-105">This article also covers workbook-level actions performed through the [Application](/javascript/api/excel/excel.application) object.</span></span>

<span data-ttu-id="78ad1-106">Объект Workbook — это точка входа для вашей надстройки для взаимодействия с Excel.</span><span class="sxs-lookup"><span data-stu-id="78ad1-106">The Workbook object is the entry point for your add-in to interact with Excel.</span></span> <span data-ttu-id="78ad1-107">Он поддерживает коллекции листов, таблиц, сводных таблиц и других элементов, через которые выполняется доступ и изменение данных Excel.</span><span class="sxs-lookup"><span data-stu-id="78ad1-107">It maintains collections of worksheets, tables, PivotTables, and more, through which Excel data is accessed and changed.</span></span> <span data-ttu-id="78ad1-108">Объект [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) предоставляет надстройке доступ ко всем данным книги с помощью отдельных листов.</span><span class="sxs-lookup"><span data-stu-id="78ad1-108">The [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) object gives your add-in access to all the workbook's data through indivual worksheets.</span></span> <span data-ttu-id="78ad1-109">В частности, он позволяет надстройке добавлять листы, перемещаться между ними и назначать обработчиков событий листа.</span><span class="sxs-lookup"><span data-stu-id="78ad1-109">Specifically, it lets your add-in add worksheets, navigate among them, and assign handlers to worksheet events.</span></span> <span data-ttu-id="78ad1-110">В статье [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md) описывается способ доступа к листам и их изменение.</span><span class="sxs-lookup"><span data-stu-id="78ad1-110">The article [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md) describes how to access and edit worksheets.</span></span>

## <a name="get-the-active-cell-or-selected-range"></a><span data-ttu-id="78ad1-111">Получение активной ячейки или выделенного диапазона</span><span class="sxs-lookup"><span data-stu-id="78ad1-111">Get the active cell or selected range</span></span>

<span data-ttu-id="78ad1-112">Объект Workbook содержит два метода для получения диапазона ячеек, выделенных пользователем или надстройкой: `getActiveCell()` и `getSelectedRange()`.</span><span class="sxs-lookup"><span data-stu-id="78ad1-112">The Workbook object contains two methods that get a range of cells the user or add-in has selected: `getActiveCell()` and `getSelectedRange()`.</span></span> <span data-ttu-id="78ad1-113">`getActiveCell()` получает активную ячейку из книги в виде [объекта Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="78ad1-113">`getActiveCell()` gets the active cell from the workbook as a [Range object](/javascript/api/excel/excel.range).</span></span> <span data-ttu-id="78ad1-114">В приведенном ниже примере показан вызов `getActiveCell()` с последующей печатью адреса ячейки в консоль.</span><span class="sxs-lookup"><span data-stu-id="78ad1-114">The following example shows a call to `getActiveCell()`, followed by the cell's address being printed to the console.</span></span>

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="78ad1-115">Метод `getSelectedRange()` возвращает один диапазон, выделенный в настоящее время.</span><span class="sxs-lookup"><span data-stu-id="78ad1-115">The `getSelectedRange()` method returns the currently selected single range.</span></span> <span data-ttu-id="78ad1-116">Если выделено несколько диапазонов, возникает ошибка InvalidSelection.</span><span class="sxs-lookup"><span data-stu-id="78ad1-116">If multiple ranges are selected, an InvalidSelection error is thrown.</span></span> <span data-ttu-id="78ad1-117">В приведенном ниже примере показан вызов метода `getSelectedRange()`, который затем устанавливает желтый цвет заливки для диапазона.</span><span class="sxs-lookup"><span data-stu-id="78ad1-117">The following example shows a call to `getSelectedRange()` that then sets the range's fill color to yellow.</span></span>

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a><span data-ttu-id="78ad1-118">Создание книги</span><span class="sxs-lookup"><span data-stu-id="78ad1-118">Create a workbook</span></span>

<span data-ttu-id="78ad1-119">Ваша надстройка может создать новую книгу, отдельную от экземпляра Excel, в котором в настоящее время работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="78ad1-119">Your add-in can create a new workbook, separate from the Excel instance in which the add-in is currently running.</span></span> <span data-ttu-id="78ad1-120">Для этой цели в объекте Excel имеется метод `createWorkbook`.</span><span class="sxs-lookup"><span data-stu-id="78ad1-120">The Excel object has the `createWorkbook` method for this purpose.</span></span> <span data-ttu-id="78ad1-121">При вызове этого метода сразу открывается и отображается новая книга в новом экземпляре программы Excel.</span><span class="sxs-lookup"><span data-stu-id="78ad1-121">When this method is called, the new workbook is immediately opened and displayed in a new instance of Excel.</span></span> <span data-ttu-id="78ad1-122">Ваша надстройка остается открытой и запущенной в предыдущей книге.</span><span class="sxs-lookup"><span data-stu-id="78ad1-122">Your add-in remains open and running with the previous workbook.</span></span>

```js
Excel.createWorkbook();
```

<span data-ttu-id="78ad1-123">С помощью метода `createWorkbook` также можно создать копию существующей книги.</span><span class="sxs-lookup"><span data-stu-id="78ad1-123">The `createWorkbook` method can also create a copy of an existing workbook.</span></span> <span data-ttu-id="78ad1-124">Метод принимает в качестве необязательного параметра строковое представление XLSX-файла в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="78ad1-124">The method accepts a base64-encoded string representation of an .xlsx file as an optional parameter.</span></span> <span data-ttu-id="78ad1-125">Полученная книга будет копией этого файла, предполагая, что строковый аргумент является допустимым XLSX-файлом.</span><span class="sxs-lookup"><span data-stu-id="78ad1-125">The resulting workbook will be a copy of that file, assuming the string argument is a valid .xlsx file.</span></span>

<span data-ttu-id="78ad1-126">Текущую книгу надстройки можно получить в виде строки в кодировке base64 с помощью [среза файла](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="78ad1-126">You can get your add-in’s current workbook as a base64-encoded string by using [file slicing](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span> <span data-ttu-id="78ad1-127">Преобразование файла в нужную строку в кодировке base64 можно выполнить с помощью класса [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader), как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="78ad1-127">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span> 

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = event.target.result.indexOf("base64,");
        var mybase64 = event.target.result.substr(startIndex + 7);

        Excel.createWorkbook(mybase64);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a><span data-ttu-id="78ad1-128">Защита структуры книги</span><span class="sxs-lookup"><span data-stu-id="78ad1-128">Protect the workbook's structure</span></span>

<span data-ttu-id="78ad1-129">Надстройка может управлять возможностью пользователя по изменению структуры книги.</span><span class="sxs-lookup"><span data-stu-id="78ad1-129">Your add-in can control a user's ability to edit the workbook's structure.</span></span> <span data-ttu-id="78ad1-130">Свойство `protection` объекта Workbook является объектом [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) с методом `protect()`.</span><span class="sxs-lookup"><span data-stu-id="78ad1-130">The Workbook object's `protection` property is a [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) object with a `protect()` method.</span></span> <span data-ttu-id="78ad1-131">В приведенном ниже примере показан основной сценарий переключения защиты структуры книги.</span><span class="sxs-lookup"><span data-stu-id="78ad1-131">The following example shows a basic scenario toggling the protection of the workbook's structure.</span></span> 

```js
Excel.run(function (context) {
    var workbook = context.workbook;
    workbook.load("protection/protected");

    return context.sync().then(function() {
        if (!workbook.protection.protected) {
            workbook.protection.protect();
        }
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="78ad1-132">Метод `protect` принимает необязательный строковый параметр.</span><span class="sxs-lookup"><span data-stu-id="78ad1-132">The `protect` method accepts an optional string parameter.</span></span> <span data-ttu-id="78ad1-133">Эта строка представляет пароль, необходимый пользователю для обхода защиты и изменения структуры книги.</span><span class="sxs-lookup"><span data-stu-id="78ad1-133">This string represents the password needed for a user to bypass protection and change the workbook's structure.</span></span>

<span data-ttu-id="78ad1-134">Защиту также можно установить на уровне книги, чтобы предотвратить нежелательные изменения данных.</span><span class="sxs-lookup"><span data-stu-id="78ad1-134">Protection can also be set at the worksheet level to prevent unwanted data editing.</span></span> <span data-ttu-id="78ad1-135">Дополнительные сведения см. в разделе **Защита данных** статьи [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md#data-protection).</span><span class="sxs-lookup"><span data-stu-id="78ad1-135">For more information, see the **Data protection** section of the [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#data-protection) article.</span></span>

> [!NOTE] 
> <span data-ttu-id="78ad1-136">Дополнительные сведения о защите книги в Excel см. в статье [Защита книги](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).</span><span class="sxs-lookup"><span data-stu-id="78ad1-136">For more information about workbook protection in Excel, see the [Protect a workbook](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517) article.</span></span>

## <a name="access-document-properties"></a><span data-ttu-id="78ad1-137">Доступ к свойствам документов</span><span class="sxs-lookup"><span data-stu-id="78ad1-137">Access document properties</span></span>

<span data-ttu-id="78ad1-138">Объекты Workbook имеют доступ к метаданным файлов Office, называемым [свойствами документов](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span><span class="sxs-lookup"><span data-stu-id="78ad1-138">Workbook objects have access to the Office file metadata, which is known as the [document properties](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75).</span></span> <span data-ttu-id="78ad1-139">Свойство `properties` объекта Workbook является объектом [DocumentProperties](/javascript/api/excel/excel.documentproperties), содержащим эти значения метаданных.</span><span class="sxs-lookup"><span data-stu-id="78ad1-139">The Workbook object's `properties` property is a [DocumentProperties](/javascript/api/excel/excel.documentproperties) object containing these metadata values.</span></span> <span data-ttu-id="78ad1-140">В приведенном ниже примере показано, как установить свойство **author**.</span><span class="sxs-lookup"><span data-stu-id="78ad1-140">The following example shows how to set the **MetadataCatalogFileName** property declaratively.</span></span>

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="78ad1-141">Также можно установить настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="78ad1-141">You can also define custom properties.</span></span> <span data-ttu-id="78ad1-142">Объект DocumentProperties содержит свойство `custom`, представляющее коллекцию пар "ключ-значение" для свойств, определяемых пользователем.</span><span class="sxs-lookup"><span data-stu-id="78ad1-142">The DocumentProperties object contains a `custom` property that represents a collection of key-value pairs for user-defined properties.</span></span> <span data-ttu-id="78ad1-143">В приведенном ниже примере показано, как создать настраиваемое свойство с именем **Introduction** со значением "Hello", а затем вызвать его.</span><span class="sxs-lookup"><span data-stu-id="78ad1-143">The following example shows how to create a custom property named **Introduction** with the value "Hello", then retrieve it.</span></span>

```js
Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    customDocProperties.add("Introduction", "Hello");
    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    var customDocProperties = context.workbook.properties.custom;
    var customProperty = customDocProperties.getItem("Introduction");
    customProperty.load("key, value");

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a><span data-ttu-id="78ad1-144">Доступ к параметрам документа</span><span class="sxs-lookup"><span data-stu-id="78ad1-144">Access document settings</span></span>

<span data-ttu-id="78ad1-145">Параметры книги похожи на коллекцию настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="78ad1-145">A workbook's settings are similar to the collection of custom properties.</span></span> <span data-ttu-id="78ad1-146">Различие заключается в том, что параметры уникальны для одного файла Excel и соединения надстройки, а свойства связаны только с файлом.</span><span class="sxs-lookup"><span data-stu-id="78ad1-146">The difference is settings are unique to a single Excel file and add-in pairing, whereas properties are solely connected to the file.</span></span> <span data-ttu-id="78ad1-147">В приведенном ниже примере показано, как создать параметр и получить к нему доступ.</span><span class="sxs-lookup"><span data-stu-id="78ad1-147">The following example shows how to create a file and add it to a folder.</span></span>

```js
Excel.run(function (context) {
    var settings = context.workbook.settings;
    settings.add("NeedsReview", true);
    var needsReview = settings.getItem("NeedsReview");
    needsReview.load("value");

    return context.sync().then(function() {
        console.log("Workbook needs review : " + needsReview.value);
    });
}).catch(errorHandlerFunction);
```

## <a name="control-calculation-behavior"></a><span data-ttu-id="78ad1-148">Управление режимом вычислений</span><span class="sxs-lookup"><span data-stu-id="78ad1-148">Control calculation behavior</span></span>

### <a name="set-calculation-mode"></a><span data-ttu-id="78ad1-149">Установка режима вычислений</span><span class="sxs-lookup"><span data-stu-id="78ad1-149">Set calculation mode</span></span>

<span data-ttu-id="78ad1-150">По умолчанию Excel пересчитывает результаты формул при каждом изменении ячейки из ссылки.</span><span class="sxs-lookup"><span data-stu-id="78ad1-150">By default, Excel recalculates formula results whenever a referenced cell is changed.</span></span> <span data-ttu-id="78ad1-151">Производительность вашей надстройки можно улучшить путем изменения режима вычислений.</span><span class="sxs-lookup"><span data-stu-id="78ad1-151">Your add-in's performance may benefit from adjusting this calculation behavior.</span></span> <span data-ttu-id="78ad1-152">У объекта Application есть свойство `calculationMode` типа `CalculationMode`.</span><span class="sxs-lookup"><span data-stu-id="78ad1-152">The Application object has a `calculationMode` property of type `CalculationMode`.</span></span> <span data-ttu-id="78ad1-153">Ему можно присвоить следующие значения:</span><span class="sxs-lookup"><span data-stu-id="78ad1-153">It can be set to the following values:</span></span>

 - <span data-ttu-id="78ad1-154">`automatic`: режим пересчета по умолчанию, при котором Excel вычисляет новые результаты формулы при каждом изменении соответствующих данных.</span><span class="sxs-lookup"><span data-stu-id="78ad1-154">`automatic`: The default recalculation behavior where Excel calculates new formula results every time the relevant data is changed.</span></span>
 - <span data-ttu-id="78ad1-155">`automaticExceptTables`: аналогично `automatic`, за исключением того, что игнорируются любые изменения значений таблиц.</span><span class="sxs-lookup"><span data-stu-id="78ad1-155">`automaticExceptTables`: Same as `automatic`, except any changes made to values in tables are ignored.</span></span>
 - <span data-ttu-id="78ad1-156">`manual`: вычисления выполняются только в том случае, если пользователь или надстройка запрашивает их.</span><span class="sxs-lookup"><span data-stu-id="78ad1-156">`manual`: Calculations only occur when the user or add-in requests them.</span></span>
 
### <a name="set-calculation-type"></a><span data-ttu-id="78ad1-157">Установка типа вычислений</span><span class="sxs-lookup"><span data-stu-id="78ad1-157">Set calculation type</span></span>

<span data-ttu-id="78ad1-158">Объект [Application](/javascript/api/excel/excel.application) предоставляет метод применения немедленного пересчета.</span><span class="sxs-lookup"><span data-stu-id="78ad1-158">The [Application](/javascript/api/excel/excel.application) object provides a method to force an immediate recalculation.</span></span> <span data-ttu-id="78ad1-159">Метод `Application.calculate(calculationType)` запускает ручной пересчет с учетом указанного типа `calculationType`.</span><span class="sxs-lookup"><span data-stu-id="78ad1-159">`Application.calculate(calculationType)` starts a manual recalculation based on the specified `calculationType`.</span></span> <span data-ttu-id="78ad1-160">Можно указать следующие значения:</span><span class="sxs-lookup"><span data-stu-id="78ad1-160">Specifies the operation to perform. The following table describes values that can be specified.</span></span>

 - <span data-ttu-id="78ad1-161">`full`: пересчет всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.</span><span class="sxs-lookup"><span data-stu-id="78ad1-161">`full`: Recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
 - <span data-ttu-id="78ad1-162">`fullRebuild`: проверка зависимых формул с последующим пересчетом всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.</span><span class="sxs-lookup"><span data-stu-id="78ad1-162">`fullRebuild`: Check dependent formulas, and then recalculate all formulas in all open workbooks, regardless of whether they have changed since the last recalculation.</span></span>
 - <span data-ttu-id="78ad1-163">`recalculate`: пересчет формул, которые были изменены (или помечены программным путем для пересчета) с момента последнего вычисления, и зависимых от них формул во всех активных книгах.</span><span class="sxs-lookup"><span data-stu-id="78ad1-163">`recalculate`: Recalculate formulas that have changed (or been programmatically marked for recalculation) since the last calculation, and formulas dependent on them, in all active workbooks.</span></span>
 
> [!NOTE] 
> <span data-ttu-id="78ad1-164">Дополнительные сведения о пересчете см. в статье [Изменение пересчета, итерации или точности формулы](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).</span><span class="sxs-lookup"><span data-stu-id="78ad1-164">For more information about recalculation, see the [Change formula recalculation, iteration, or precision](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4) article.</span></span>

### <a name="temporarily-suspend-calculations"></a><span data-ttu-id="78ad1-165">Временная приостановка вычисления</span><span class="sxs-lookup"><span data-stu-id="78ad1-165">Temporarily suspend calculations</span></span>

<span data-ttu-id="78ad1-166">API Excel также позволяет надстройкам отключить вычисления до вызова `RequestContext.sync()`.</span><span class="sxs-lookup"><span data-stu-id="78ad1-166">The Excel API also lets add-ins turn off calculations until `RequestContext.sync()` is called.</span></span> <span data-ttu-id="78ad1-167">Для этого используется `suspendApiCalculationUntilNextSync()`.</span><span class="sxs-lookup"><span data-stu-id="78ad1-167">This is done with `suspendApiCalculationUntilNextSync()`.</span></span> <span data-ttu-id="78ad1-168">Используйте этот метод, если ваша надстройка изменяет большие диапазоны без необходимости доступа к данным между изменениями.</span><span class="sxs-lookup"><span data-stu-id="78ad1-168">Use this method when your add-in is editing large ranges without needing to access the data between edits.</span></span>

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="see-also"></a><span data-ttu-id="78ad1-169">См. также</span><span class="sxs-lookup"><span data-stu-id="78ad1-169">See also</span></span>

- [<span data-ttu-id="78ad1-170">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="78ad1-170">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="78ad1-171">Работа с листами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="78ad1-171">Work with Worksheets using the Excel JavaScript API</span></span>](excel-add-ins-worksheets.md)
- [<span data-ttu-id="78ad1-172">Работа с диапазонами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="78ad1-172">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)