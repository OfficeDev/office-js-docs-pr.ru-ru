---
title: Работа с книгами с использованием API JavaScript для Excel
description: Примеры кода, в которых показано, как выполнять распространенные задачи с книгами или функциями уровня приложения с помощью API JavaScript для Excel.
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: a7a35e2627863c648f8c3e31ab05b2714ca0aebe
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294131"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>Работа с книгами с использованием API JavaScript для Excel

В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для книг с использованием API JavaScript для Excel. Полный список свойств и методов, `Workbook` поддерживаемых объектом, представлен в статье [объект WORKBOOK (API JavaScript для Excel)](/javascript/api/excel/excel.workbook). В этой статье также рассматриваются действия на уровне книги, выполняемые с помощью объекта [Application](/javascript/api/excel/excel.application).

Объект Workbook — это точка входа для вашей надстройки для взаимодействия с Excel. Он поддерживает коллекции листов, таблиц, сводных таблиц и других элементов, через которые выполняется доступ и изменение данных Excel. Объект [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) предоставляет надстройке доступ ко всем данным книги с помощью отдельных листов. В частности, он позволяет надстройке добавлять листы, перемещаться между ними и назначать обработчиков событий листа. В статье [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md) описывается способ доступа к листам и их изменение.

## <a name="get-the-active-cell-or-selected-range"></a>Получение активной ячейки или выделенного диапазона

Объект Workbook содержит два метода для получения диапазона ячеек, выделенных пользователем или надстройкой: `getActiveCell()` и `getSelectedRange()`. `getActiveCell()` получает активную ячейку из книги в виде [объекта Range](/javascript/api/excel/excel.range). В приведенном ниже примере показан вызов `getActiveCell()` с последующей печатью адреса ячейки в консоль.

```js
Excel.run(function (context) {
    var activeCell = context.workbook.getActiveCell();
    activeCell.load("address");

    return context.sync().then(function () {
        console.log("The active cell is " + activeCell.address);
    });
}).catch(errorHandlerFunction);
```

Метод `getSelectedRange()` возвращает один диапазон, выделенный в настоящее время. Если выделено несколько диапазонов, возникает ошибка InvalidSelection. В приведенном ниже примере показан вызов метода `getSelectedRange()`, который затем устанавливает желтый цвет заливки для диапазона.

```js
Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    range.format.fill.color = "yellow";
    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="create-a-workbook"></a>Создание книги

Ваша надстройка может создать новую книгу, отдельную от экземпляра Excel, в котором в настоящее время работает надстройка. Для этой цели в объекте Excel имеется метод `createWorkbook`. При вызове этого метода сразу открывается и отображается новая книга в новом экземпляре программы Excel. Ваша надстройка остается открытой и запущенной в предыдущей книге.

```js
Excel.createWorkbook();
```

С помощью метода `createWorkbook` также можно создать копию существующей книги. Метод принимает в качестве необязательного параметра строковое представление XLSX-файла в кодировке base64. Полученная книга будет копией этого файла, предполагая, что строковый аргумент является допустимым XLSX-файлом.

Вы можете получить текущую книгу надстройки в виде строки в кодировке Base64 с помощью [фрагментирования файлов](/javascript/api/office/office.document#getfileasync-filetype--options--callback-). Преобразование файла в нужную строку в кодировке base64 можно выполнить с помощью класса [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader), как показано в приведенном ниже примере.

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(workbookContents);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one-preview"></a>Вставьте копию существующей книги в текущую (предварительная версия)

> [!NOTE]
> В настоящее время метод `WorksheetCollection.addFromBase64` доступен только в общедоступной предварительной версии и только в Office для Windows и Mac. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

В предыдущем примере показана новая книга, которая была создана из существующей книги. Вы также можете скопировать отдельные части или всю существующую книгу целиком в книгу, привязанную в настоящее время к вашей надстройке. [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) для книги имеет метод `addFromBase64` для вставки копий листов целевой книги в саму книгу. Файл другой книги передается в виде строки в кодировке base64, как и вызов `Excel.createWorkbook`.

```TypeScript
addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet | string): OfficeExtension.ClientResult<string[]>;
```

В примере ниже показаны листы книги, которые были вставлены в текущую книгу непосредственно после активного листа. Обратите внимание, что `null` передается для параметра `sheetNamesToInsert?: string[]`. Это означает, что все листы были вставлены.

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // strip off the metadata before the base64-encoded string
        var startIndex = reader.result.toString().indexOf("base64,");
        var workbookContents = reader.result.toString().substr(startIndex + 7);

        var sheets = context.workbook.worksheets;
        sheets.addFromBase64(
            workbookContents,
            null, // get all the worksheets
            Excel.WorksheetPositionType.after, // insert them after the worksheet specified by the next parameter
            sheets.getActiveWorksheet() // insert them after the active worksheet
        );
        return context.sync();
    });
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="protect-the-workbooks-structure"></a>Защита структуры книги

Надстройка может управлять возможностью пользователя по изменению структуры книги. Свойство `protection` объекта Workbook является объектом [WorkbookProtection](/javascript/api/excel/excel.workbookprotection) с методом `protect()`. В приведенном ниже примере показан основной сценарий переключения защиты структуры книги.

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

Метод `protect` принимает необязательный строковый параметр. Эта строка представляет пароль, необходимый пользователю для обхода защиты и изменения структуры книги.

Защиту также можно установить на уровне книги, чтобы предотвратить нежелательные изменения данных. Дополнительные сведения см. в разделе **Защита данных** статьи [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md#data-protection).

> [!NOTE]
> Дополнительные сведения о защите книги в Excel см. в статье [Защита книги](https://support.office.com/article/Protect-a-workbook-7E365A4D-3E89-4616-84CA-1931257C1517).

## <a name="access-document-properties"></a>Доступ к свойствам документов

Объекты Workbook имеют доступ к метаданным файлов Office, называемым [свойствами документов](https://support.office.com/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75). Свойство `properties` объекта Workbook является объектом [DocumentProperties](/javascript/api/excel/excel.documentproperties), содержащим эти значения метаданных. В приведенном ниже примере показано, как задать `author` свойство.

```js
Excel.run(function (context) {
    var docProperties = context.workbook.properties;
    docProperties.author = "Alex";
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="custom-properties"></a>Настраиваемые свойства

Также можно установить настраиваемые свойства. Объект DocumentProperties содержит свойство `custom`, представляющее коллекцию пар "ключ-значение" для свойств, определяемых пользователем. В приведенном ниже примере показано, как создать настраиваемое свойство с именем **Introduction** со значением "Hello", а затем вызвать его.

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
    customProperty.load(["key, value"]);

    return context.sync().then(function() {
        console.log("Custom key  : " + customProperty.key); // "Introduction"
        console.log("Custom value : " + customProperty.value); // "Hello"
    });
}).catch(errorHandlerFunction);
```

#### <a name="worksheet-level-custom-properties-preview"></a>Настраиваемые свойства на уровне листа (Предварительная версия)

> [!NOTE]
> Настраиваемые свойства на уровне листа в настоящее время находятся в режиме предварительного просмотра. [!INCLUDE [Information about using preview Excel APIs](../includes/using-excel-preview-apis.md)]

Настраиваемые свойства также можно задать на уровне листа. Они похожи на настраиваемые свойства на уровне документа, за исключением того, что один и тот же ключ может повторяться на разных листах. В приведенном ниже примере показано, как создать настраиваемое свойство с именем **воркшитграуп** со значением "Alpha" на текущем листе, а затем извлечь его.

```js
Excel.run(function (context) {
    // Add the custom property.
    var customWorksheetProperties = context.workbook.worksheets.getActiveWorksheet().customProperties;
    customWorksheetProperties.add("WorksheetGroup", "Alpha");

    return context.sync();
}).catch(errorHandlerFunction);

[...]

Excel.run(function (context) {
    // Load the keys and values of all custom properties in the current worksheet.
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.load("name");

    var customWorksheetProperties = worksheet.customProperties;
    var customWorksheetProperty = customWorksheetProperties.getItem("WorksheetGroup");
    customWorksheetProperty.load(["key", "value"]);

    return context.sync().then(function() {
        // Log the WorksheetGroup custom property to the console.
        console.log(worksheet.name + ": " + customWorksheetProperty.key); // "WorksheetGroup"
        console.log("  Custom value : " + customWorksheetProperty.value); // "Alpha"
    });
}).catch(errorHandlerFunction);
```

## <a name="access-document-settings"></a>Доступ к параметрам документа

Параметры книги похожи на коллекцию настраиваемых свойств. Различие заключается в том, что параметры уникальны для одного файла Excel и соединения надстройки, а свойства связаны только с файлом. В приведенном ниже примере показано, как создать параметр и получить к нему доступ.

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

## <a name="access-application-culture-settings"></a>Параметры культуры приложения Access

Книга содержит параметры языка и региональных параметров, которые влияют на отображение определенных данных. Эти параметры могут помочь локализовать данные, когда пользователи надстройки совместно работают с книгами на различных языках и региональных параметрах. Надстройка может использовать синтаксический анализ строк для локализации формата чисел, дат и времени на основе параметров языковых параметров системы, чтобы каждый пользователь видел данные в формате языка и региональных параметров.

`Application.cultureInfo` Определяет параметры языка и региональных параметров системы в виде объекта [CultureInfo](/javascript/api/excel/excel.cultureinfo) . Содержит такие параметры, как числовой десятичный разделитель или формат даты.

Некоторые параметры культуры можно [изменить с помощью пользовательского интерфейса Excel](https://support.office.com/article/Change-the-character-used-to-separate-thousands-or-decimals-c093b545-71cb-4903-b205-aebb9837bd1e). Параметры системы сохраняются в `CultureInfo` объекте. Все локальные изменения хранятся в виде свойств уровня [приложения](/javascript/api/excel/excel.application), например `Application.decimalSeparator` .

В примере ниже показано, как изменить символ десятичного разделителя в числовой строке с "," на символ, используемый параметрами системы.

```js
// This will convert a number like "14,37" to "14.37"
// (assuming the system decimal separator is ".").
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var decimalSource = sheet.getRange("B2");
    decimalSource.load("values");
    context.application.cultureInfo.numberFormat.load("numberDecimalSeparator");

    return context.sync().then(function() {
        var systemDecimalSeparator =
            context.application.cultureInfo.numberFormat.numberDecimalSeparator;
        var oldDecimalString = decimalSource.values[0][0];

        // This assumes the input column is standardized to use "," as the decimal separator.
        var newDecimalString = oldDecimalString.replace(",", systemDecimalSeparator);

        var resultRange = sheet.getRange("C2");
        resultRange.values = [[newDecimalString]];
        resultRange.format.autofitColumns();
        return context.sync();
    });
});
```

## <a name="add-custom-xml-data-to-the-workbook"></a>Добавление настраиваемых XML-данных в книгу

Формат файла Excel Open XML **(XLSX)** позволяет надстройке внедрить настраиваемые XML-данные в книгу. Эти данные сохраняются с книгой независимо от надстройки.

Книга содержит объект [CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection), являющийся списком объектов [CustomXmlParts](/javascript/api/excel/excel.customxmlpart). Они предоставляют доступ к строкам XML и соответствующему уникальному идентификатору. Сохраняя эти идентификаторы как параметры, надстройка может сохранять ключи к частям XML между сеансами.

В приведенных ниже примерах показано, как использовать настраиваемые части XML. В первом блоке кода показано, как внедрять XML-данные в документ. Выполняется сохранение списка проверяющих, а затем используются параметры книги, чтобы сохранить параметр `id` XML для будущих извлечений. Во втором блоке показано, как получить доступ к этим XML-данным позднее. Параметр "ContosoReviewXmlPartId" загружается и передается объекту `customXmlParts` книги. Данные XML затем печатаются в консоль.

```js
Excel.run(async (context) => {
    // Add reviewer data to the document as XML
    var originalXml = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    var customXmlPart = context.workbook.customXmlParts.add(originalXml);
    customXmlPart.load("id");

    return context.sync().then(function() {
        // Store the XML part's ID in a setting
        var settings = context.workbook.settings;
        settings.add("ContosoReviewXmlPartId", customXmlPart.id);
    });
}).catch(errorHandlerFunction);
```

```js
Excel.run(async (context) => {
    // Retrieve the XML part's id from the setting
    var settings = context.workbook.settings;
    var xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    return context.sync().then(function () {
        if (xmlPartIDSetting.value) {
            var customXmlPart = context.workbook.customXmlParts.getItem(xmlPartIDSetting.value);
            var xmlBlob = customXmlPart.getXml();

            return context.sync().then(function () {
                // Add spaces to make more human readable in the console
                var readableXML = xmlBlob.value.replace(/></g, "> <");
                console.log(readableXML);
            });
        }
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> `CustomXMLPart.namespaceUri` заполняется только в том случае, если настраиваемый XML-элемент верхнего уровня содержит атрибут `xmlns`.

## <a name="control-calculation-behavior"></a>Управление режимом вычислений

### <a name="set-calculation-mode"></a>Установка режима вычислений

По умолчанию Excel пересчитывает результаты формул при каждом изменении ячейки из ссылки. Производительность вашей надстройки можно улучшить путем изменения режима вычислений. У объекта Application есть свойство `calculationMode` типа `CalculationMode`. Ему можно присвоить следующие значения:

- `automatic`: режим пересчета по умолчанию, при котором Excel вычисляет новые результаты формулы при каждом изменении соответствующих данных.
- `automaticExceptTables`: аналогично `automatic`, за исключением того, что игнорируются любые изменения значений таблиц.
- `manual`: вычисления выполняются только в том случае, если пользователь или надстройка запрашивает их.

### <a name="set-calculation-type"></a>Установка типа вычислений

Объект [Application](/javascript/api/excel/excel.application) предоставляет метод применения немедленного пересчета. Метод `Application.calculate(calculationType)` запускает ручной пересчет с учетом указанного типа `calculationType`. Можно указать следующие значения:

- `full`: пересчет всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.
- `fullRebuild`: проверка зависимых формул с последующим пересчетом всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.
- `recalculate`: пересчет формул, которые были изменены (или помечены программным путем для пересчета) с момента последнего вычисления, и зависимых от них формул во всех активных книгах.

> [!NOTE]
> Дополнительные сведения о пересчете см. в статье [Изменение пересчета, итерации или точности формулы](https://support.office.com/article/change-formula-recalculation-iteration-or-precision-73fc7dac-91cf-4d36-86e8-67124f6bcce4).

### <a name="temporarily-suspend-calculations"></a>Временная приостановка вычисления

API Excel также позволяет надстройкам отключить вычисления до вызова `RequestContext.sync()`. Для этого используется `suspendApiCalculationUntilNextSync()`. Используйте этот метод, если ваша надстройка изменяет большие диапазоны без необходимости доступа к данным между изменениями.

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="save-the-workbook"></a>Сохраните книгу.

`Workbook.save` сохраняет книгу в постоянное хранилище. Метод `save` имеет один необязательный параметр `saveBehavior`, который может принимать одно из следующих значений:

- `Excel.SaveBehavior.save` (по умолчанию): файл будет сохранен без предварительного запроса имени файла, а также место для сохранения. Если файл не был сохранен ранее, он будет сохранен в папке по умолчанию. Если файл уже был сохранен ранее, он будет сохранен в той же папке.
- `Excel.SaveBehavior.prompt`: если файл не был сохранен ранее, будет предложено ввести имя файла и место для сохранения. Если файл уже был сохранен ранее, он будет сохраняться в той же папке, и никаких дополнительных действий не потребуется.

> [!CAUTION]
> Если пользователь при запрос на сохранение отменяет операцию, `save` выдает исключение.

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a>Закрытие книги.

`Workbook.close` закрывает книгу, а также надстройки, связанные с книгой, (приложение Excel остается открытым). Метод `close` имеет один необязательный параметр `closeBehavior`, который может принимать одно из следующих значений:

- `Excel.CloseBehavior.save` (по умолчанию): файл будет сохранен до закрытия. Если файл не был сохранен ранее, будет предложено ввести имя файла и место для сохранения.
- `Excel.CloseBehavior.skipSave`: файл будет немедленно закрыт без сохранения. Все несохраненные изменения будут потеряны.

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md)
- [Работа с диапазонами с использованием API JavaScript для Excel](excel-add-ins-ranges.md)
