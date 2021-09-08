---
title: Работа с книгами с использованием API JavaScript для Excel
description: Узнайте, как выполнять общие задачи с помощью книг или функций уровня приложений с помощью Excel API JavaScript.
ms.date: 06/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ed63337aad322762019e8a51e3f1cc1c202db210
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936671"
---
# <a name="work-with-workbooks-using-the-excel-javascript-api"></a>Работа с книгами с использованием API JavaScript для Excel

В этой статье приведены примеры кода, в которых показано, как выполнять стандартные задачи для книг с использованием API JavaScript для Excel. Полный список свойств и методов, поддерживаемых объектом, см. в книге `Workbook` [Объект (API JavaScript для Excel).](/javascript/api/excel/excel.workbook) В этой статье также рассматриваются действия на уровне книги, выполняемые с помощью объекта [Application](/javascript/api/excel/excel.application).

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

Текущую книгу надстройки можно получить в качестве строки с кодом base64 с помощью [нарезки файлов.](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_) Преобразование файла в нужную строку в кодировке base64 можно выполнить с помощью класса [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader), как показано в приведенном ниже примере.

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (function (event) {
    Excel.run(function (context) {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);

        Excel.createWorkbook(externalWorkbook);
        return context.sync();
    }).catch(errorHandlerFunction);
});

// Read the file as a data URL so we can parse the base64-encoded string.
reader.readAsDataURL(myFile.files[0]);
```

### <a name="insert-a-copy-of-an-existing-workbook-into-the-current-one"></a>Вставка копии существующей книги в текущую книгу.

В предыдущем примере показана новая книга, которая была создана из существующей книги. Вы также можете скопировать отдельные части или всю существующую книгу целиком в книгу, привязанную в настоящее время к вашей надстройке. В [книге](/javascript/api/excel/excel.workbook) используется метод вставки копий таблиц целевой книги `insertWorksheetsFromBase64` в себя. Файл другой книги передается как строка с кодом base64, как и `Excel.createWorkbook` вызов. 

```TypeScript
insertWorksheetsFromBase64(base64File: string, options?: Excel.InsertWorksheetOptions): OfficeExtension.ClientResult<string[]>;
```

> [!IMPORTANT]
> Метод `insertWorksheetsFromBase64` поддерживается для Excel на Windows, Mac и в Интернете. Он не поддерживается для iOS. Кроме того, Excel в Интернете этот метод не поддерживает исходные таблицы с элементами PivotTable, Chart, Comment или Slicer. Если эти объекты присутствуют, метод возвращает `insertWorksheetsFromBase64` `UnsupportedFeature` ошибку в Excel в Интернете. 

В следующем примере кода показано, как вставить в текущую книгу таблицы из другой книги. Этот пример кода сначала обрабатывает файл книги с объектом и извлекает строку с кодом base64, а затем вставляет эту строку с кодом base64 в текущую [`FileReader`](https://developer.mozilla.org/docs/Web/API/FileReader) книгу. Новые листы вставляются после листа с именем **Sheet1.** Обратите внимание, что он передается в качестве параметра свойства `[]` [InsertWorksheetOptions.sheetNamesToInsert.](/javascript/api/excel/excel.insertworksheetoptions#sheetNamesToInsert) Это означает, что все таблицы из целевой книги вставляются в текущую книгу.

```js
// Retrieve the external workbook file and set up a `FileReader` object. 
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = (event) => {
    Excel.run((context) => {
        // Remove the metadata before the base64-encoded string.
        var startIndex = reader.result.toString().indexOf("base64,");
        var externalWorkbook = reader.result.toString().substr(startIndex + 7);
            
        // Retrieve the current workbook.
        var workbook = context.workbook;
            
        // Set up the insert options. 
        var options = { 
            sheetNamesToInsert: [], // Insert all the worksheets from the source workbook.
            positionType: Excel.WorksheetPositionType.after, // Insert after the `relativeTo` sheet.
            relativeTo: "Sheet1" // The sheet relative to which the other worksheets will be inserted. Used with `positionType`.
        }; 
            
         // Insert the new worksheets into the current workbook.
         workbook.insertWorksheetsFromBase64(externalWorkbook, options);
         return context.sync();
    });
};

// Read the file as a data URL so we can parse the base64-encoded string.
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
> Дополнительные сведения о защите книги в Excel см. в статье [Защита книги](https://support.microsoft.com/office/7e365a4d-3e89-4616-84ca-1931257c1517).

## <a name="access-document-properties"></a>Доступ к свойствам документов

Объекты Workbook имеют доступ к метаданным файлов Office, называемым [свойствами документов](https://support.microsoft.com/office/21d604c2-481e-4379-8e54-1dd4622c6b75). Свойство `properties` объекта Workbook является объектом [DocumentProperties](/javascript/api/excel/excel.documentproperties), содержащим эти значения метаданных. В следующем примере показано, как установить `author` свойство.

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

#### <a name="worksheet-level-custom-properties"></a>Настраиваемые свойства на уровне таблицы

Настраиваемые свойства также можно установить на уровне таблицы. Они похожи на настраиваемые свойства на уровне документов, за исключением того, что один и тот же ключ может повторяться в разных таблицах. В следующем примере показано, как создать настраиваемую свойство **WorksheetGroup** со значением "Альфа" на текущем таблице, а затем получить его.

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

## <a name="access-application-culture-settings"></a>Доступ к настройкам культуры приложений

В книге есть языковые и культурные параметры, влияющие на отображение определенных данных. Эти параметры могут помочь локализовать данные, когда пользователи надстройки делятся книгами на разных языках и культурах. Ваша надстройка может использовать анализ строк для локализации формата чисел, дат и времени в зависимости от параметров культуры системы, чтобы каждый пользователь видел данные в формате своей культуры.

`Application.cultureInfo`определяет параметры культуры системы как объект [CultureInfo.](/javascript/api/excel/excel.cultureinfo) Это содержит параметры, такие как числовой десятичной сепаратор или формат даты.

Некоторые параметры культуры можно [изменить с помощью Excel пользовательского интерфейса.](https://support.microsoft.com/office/c093b545-71cb-4903-b205-aebb9837bd1e) Параметры системы сохраняются в `CultureInfo` объекте. Любые локальные изменения хранятся как [свойства уровня приложения,](/javascript/api/excel/excel.application)например `Application.decimalSeparator` .

В следующем примере изменяется десятичное сепараторное течение числовой строки с "," на символ, используемый в параметрах системы.

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

По умолчанию Excel пересчитывает результаты формул при каждом изменении ячейки из ссылки. Производительность вашей надстройки можно улучшить путем изменения режима вычислений. У объекта Application есть свойство `calculationMode` типа `CalculationMode`. Его можно установить к следующим значениям.

- `automatic`: режим пересчета по умолчанию, при котором Excel вычисляет новые результаты формулы при каждом изменении соответствующих данных.
- `automaticExceptTables`: аналогично `automatic`, за исключением того, что игнорируются любые изменения значений таблиц.
- `manual`: вычисления выполняются только в том случае, если пользователь или надстройка запрашивает их.

### <a name="set-calculation-type"></a>Установка типа вычислений

Объект [Application](/javascript/api/excel/excel.application) предоставляет метод применения немедленного пересчета. Метод `Application.calculate(calculationType)` запускает ручной пересчет с учетом указанного типа `calculationType`. Можно укаварить следующие значения.

- `full`: пересчет всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.
- `fullRebuild`: проверка зависимых формул с последующим пересчетом всех формул во всех открытых книгах независимо от их изменения с прошлого пересчета.
- `recalculate`: пересчет формул, которые были изменены (или помечены программным путем для пересчета) с момента последнего вычисления, и зависимых от них формул во всех активных книгах.

> [!NOTE]
> Дополнительные сведения о пересчете см. в статье [Изменение пересчета, итерации или точности формулы](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4).

### <a name="temporarily-suspend-calculations"></a>Временная приостановка вычисления

API Excel также позволяет надстройкам отключить вычисления до вызова `RequestContext.sync()`. Для этого используется `suspendApiCalculationUntilNextSync()`. Используйте этот метод, если ваша надстройка изменяет большие диапазоны без необходимости доступа к данным между изменениями.

```js
context.application.suspendApiCalculationUntilNextSync();
```

## <a name="detect-workbook-activation"></a>Обнаружение активации книг

Ваша надстройка может обнаруживать при активации книги. Книга становится  неактивной, когда пользователь переключает фокус на другую книгу, на другое приложение или (в Excel в Интернете) на другую вкладку веб-браузера. Книга *активируется,* когда пользователь возвращает фокус в книгу. Активация книги может вызвать функции вызова в надстройке, например освежающие данные книги.

Чтобы определить, когда книга активирована, [зарегистрируйте](excel-add-ins-events.md#register-an-event-handler) обработник событий для [события onActivated](/javascript/api/excel/excel.workbook#onActivated) книги. Обработчики событий `onActivated` для события получают объект [WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs) при пожаре события.

> [!IMPORTANT]
> Событие `onActivated` не определяет, когда книга открывается. Это событие обнаруживает только тогда, когда пользователь переключается на уже открытую книгу.

В следующем примере кода показано, как зарегистрировать обработник событий `onActivated` и настроить функцию вызова.

```js
Excel.run(function (context) {
    // Retrieve the workbook.
    var workbook = context.workbook;

    // Register the workbook activated event handler.
    workbook.onActivated.add(workbookActivated);

    return context.sync();
});

function workbookActivated(event) {
    Excel.run(function (context) {
        // Retrieve the workbook and load the name.
        var workbook = context.workbook;
        workbook.load("name");
        
        return context.sync().then(function () {
            // Callback function for when the workbook is activated.
            console.log(`The workbook ${workbook.name} was activated.`);
        });
    });
}
```

## <a name="save-the-workbook"></a>Сохраните книгу.

`Workbook.save` сохраняет книгу в постоянное хранилище. Метод `save` принимает один необязательный `saveBehavior` параметр, который может быть одним из следующих значений.

- `Excel.SaveBehavior.save` (по умолчанию): файл будет сохранен без предварительного запроса имени файла, а также место для сохранения. Если файл не был сохранен ранее, он будет сохранен в папке по умолчанию. Если файл уже был сохранен ранее, он будет сохранен в той же папке.
- `Excel.SaveBehavior.prompt`: если файл не был сохранен ранее, будет предложено ввести имя файла и место для сохранения. Если файл уже был сохранен ранее, он будет сохраняться в той же папке, и никаких дополнительных действий не потребуется.

> [!CAUTION]
> Если пользователь при запрос на сохранение отменяет операцию, `save` выдает исключение.

```js
context.workbook.save(Excel.SaveBehavior.prompt);
```

## <a name="close-the-workbook"></a>Закрытие книги.

`Workbook.close` закрывает книгу, а также надстройки, связанные с книгой, (приложение Excel остается открытым). Метод `close` принимает один необязательный `closeBehavior` параметр, который может быть одним из следующих значений.

- `Excel.CloseBehavior.save` (по умолчанию): файл будет сохранен до закрытия. Если файл не был сохранен ранее, будет предложено ввести имя файла и место для сохранения.
- `Excel.CloseBehavior.skipSave`: файл будет немедленно закрыт без сохранения. Все несохраненные изменения будут потеряны.

```js
context.workbook.close(Excel.CloseBehavior.save);
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с листами с использованием API JavaScript для Excel](excel-add-ins-worksheets.md)
