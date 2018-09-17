---
title: Работа с несколькими диапазонами одновременно в надстройках Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: ade97947e513d0af5d7a520c1f07ef1fa046dd0f
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23949853"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Работа с несколькими диапазонами одновременно в надстройках Excel (предварительная версия)

Библиотека Excel JavaScript позволяет вашей надстройке выполнять операции и устанавливать свойства одновременно для нескольких диапазонов. Диапазоны не должны быть непрерывными. В дополнение к упрощению вашего кода этот способ установки свойства выполняется намного быстрее, чем установка этого же свойства индивидуально для каждого из диапазонов.

> [!NOTE]
> Для API-интерфейсов, описанных в этой статье, требуется **версия Office 2016 Click-to-Run 1809 сборки 10820.20000** или более поздняя версия. (Возможно, вам потребуется присоединиться к [программе предварительной оценки Office](https://products.office.com/office-insider) для получения соответствующей сборки.) Кроме того, необходимо загрузить бета-версию библиотеки Office JavaScript из [сети CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). К тому же, у нас еще нет страниц со ссылкой на эти API. Но следующий файл типа определения содержит описания для них: [office.d.ts бета-версии](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).

## <a name="rangeareas"></a>RangeAreas

Набор диапазонов (возможно, разобщенных) представлен объектом `Excel.RangeAreas`. Он имеет свойства и методы, аналогичные типу `Range` (многие из которых имеют одинаковые или похожие имена), но изменения были внесены в:

- Типы данных для свойств и поведений методов задания и методов получения.
- Типы данных параметров метода и поведений метода.
- Типы данных возвращаемых значений метода.

Некоторые примеры:

- `RangeAreas` имеет свойство `address`, которое возвращает строку с адресами диапазона, разделенными диапазонами, а не только один адрес, как в случае со свойством `Range.address`.
- `RangeAreas` имеет свойство `dataValidation`, которое возвращает объект`DataValidation`, представляющий проверку данных всех диапазонов в `RangeAreas`при соответствии. Свойство является`null`, если идентичные объекты `DataValidation` не применяются к каждому из всех диапазонов в `RangeAreas`. Это общие, но не универсальные принципы для объекта`RangeAreas`: *если свойство не имеет согласованных значений для каждого из всех диапазонов в `RangeAreas`, тогда оно является `null`.* См. [свойства чтения RangeAreas](#reading-properties-of-rangeareas), чтобы ознакомиться с дополнительными сведениями и некоторыми исключениями.
- `RangeAreas.cellCount` возвращает общее число ячеек во все диапазоны в `RangeAreas`.
- `RangeAreas.calculate` пересчитывает ячейки всех диапазонов в `RangeAreas`.
- `RangeAreas.getEntireColumn` и `RangeAreas.getEntireRow` возвращают другой объект `RangeAreas`, представляющий все столбцы (или строки) во всех диапазонах в `RangeAreas`. Например, если `RangeAreas` представляет "A1:C4" и "F14:L15", то `RangeAreas.getEntireColumn` возвращает объект `RangeAreas`, представляющий "A:C" и "F:L".
- `RangeAreas.copyFrom` можно использовать параметр `Range` или `RangeAreas`, представляющий диапазон(ы) источника операции копирования.

### <a name="rangearea-specific-properties-and-methods"></a>Свойства и методы, характерные для объекта RangeArea

Тип `RangeAreas` имеет некоторые свойства и методы, которые не входят в объект `Range`:

- `areas`объект `RangeCollection`, содержащий все диапазоны, представленные объектом `RangeAreas`. Объект `RangeCollection` – также новый и аналогичен другим объектам коллекции Excel. Он имеет свойство `items`, которое представляет собой массив из объектов `Range`, представляющих диапазоны.
- `areaCount`: Общее число диапазонов в `RangeAreas`.
- `getOffsetRangeAreas`: Работает так же, как [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), за исключением того, что `RangeAreas` возвращается и содержит диапазоны, каждый из которых смещен от одного из диапазонов в исходном `RangeAreas`.

## <a name="create-rangeareas-and-set-properties"></a>Создание RangeAreas и установка свойств

Можно создать объект `RangeAreas` двумя основными способами:

- Вызовите `Worksheet.getRanges()` и передайте его в строку с адресами диапазона, разделенными запятыми. Если диапазон, который вы хотите включить, был переделан в [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), вы можете включить в строку имя вместо адреса.
- Вызовите `Workbook.getSelectedRanges()`. Этот метод возвращает `RangeAreas`, представляющий все диапазоны, выбранные на активном в данный момент листе.

После получения объекта `RangeAreas` можно создать другие с помощью методов, применяемых к объекту, который возвращает `RangeAreas`, такие как `getOffsetRangeAreas` и `getIntersection`.

> [!NOTE]
> Нельзя непосредственно добавить дополнительные диапазоны для объекта `RangeAreas`. Например, коллекция в `RangeAreas.areas` не имеет метода `add`.


> [!WARNING] 
> Не пытайтесь непосредственно добавлять или удалять элементы из массива `RangeAreas.areas.items`. Это приведет к нежелательному поведению в вашем коде. К примеру, существует возможность принудительно добавить дополнительный объект `Range` в массив, но это приведет к ошибкам, так как свойства и методы `RangeAreas` функционируют так, как если бы новый элемент не был добавлен. Например, свойство `areaCount` не включает диапазоны, принудительно добавленные таким образом, а `RangeAreas.getItemAt(index)` вызывает ошибку, если `index` больше, чем `areasCount-1`. Аналогично, удаление объекта `Range` в диапазоне `RangeAreas.areas.items` путем получения ссылки на него и вызов его метода `Range.delete` вызывает ошибки: хотя объект `Range`*будет* удален, свойства и методы родительского объекта `RangeAreas` будут вести себя так, как если бы он все еще присутствовал (или будут стремиться к таком поведению). Например, если код вызывает метод `RangeAreas.calculate`, Office будет пытаться рассчитать диапазон, но это завершится ошибкой, поскольку отсутствует объект range.

Установка свойства для `RangeAreas` задает соответствующее свойство для всех диапазонов в коллекции `RangeAreas.areas`.

Ниже приведен пример установки свойства для нескольких диапазонов. Функция выделяет диапазоны **F3:F5** и **H3:H5**.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Этот пример применяется к сценариям, в которых можно создать серьезный код адресов диапазона, передаваемых в `getRanges`, или легко рассчитать их во время выполнения. Ниже перечислены некоторые сценарии, в которых это возможно: 

- Код выполняется в контексте известного шаблона.
- Код выполняется в контексте импортированных данных, в котором известна схема данных.

Когда во время создания кода не известно, с какими диапазонами вам придется работать, необходимо обнаружить их во время выполнения. В следующем разделе описываются эти сценарии.

### <a name="discover-range-areas-programmatically"></a>Обнаружение областей диапазона с помощью программных средств

Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` можно использовать для поиска во время выполнения диапазонов, с которыми вы хотите работать, на основе характеристик ячеек и типа значений в ячейках. Вот подписи методов из файла типов данных TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

Ниже приведен пример использования первого из них. Вот что нужно знать об этом коде:

- Он ограничивает часть листа, которую нужно искать, вызвав сначала `Worksheet.getUsedRange`, а затем вызвав `getSpecialCells` только для этого диапазона.
- Передает в качестве параметра для `getSpecialCells` версию строки значения из перечисления `Excel.SpecialCellType`. Некоторые другие значения, которые могут быть переданы вместо этого, – это "Blanks" для пустых ячеек, "Constants" для ячейки со значениями литералов вместо формул и "SameConditionalFormat" для ячеек, у которых такое же условное форматирование, как и у первой ячейки в `usedRange`. Первая ячейка является верхней крайней слева ячейкой. Полный список значений перечисления см. в [office.d.ts бета-версии](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).
- Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами залиты розовым цветом даже в том случае, если не все они непрерывны. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

Иногда не обнаруживается *никаких* ячеек с целевой характеристикой. Если `getSpecialCells` не находит требуемой ячейки, он вызывает ошибку **ItemNotFound**. Это будет переадресовать поток управления к блоку или методу `catch`, если он существует. Если нет, ошибка будет останавливать функцию. Могут быть сценарии, в которых выдача ошибки – это именно то, что должно происходить при отсутствуют ячейки с целевой характеристикой. 

Однако в сценариях, для которых это нормально, но, возможно, необычно, может не оказаться соответствующих ячеек. Ваш код должен проверить наличие такой возможности и аккуратно провести работу с сценарием без выдачи ошибки. Для этих сценариев следует использовать метод `getSpecialCellsOrNullObject` и протестировать свойство `RangeAreas.isNullObject`. Ниже приведен пример. Вот что нужно знать об этом коде:

- Метод `getSpecialCellsOrNullObject` всегда возвращает объект прокси-сервера, поэтому он не может быть `null` в обычном смысле JavaScript. Но если не обнаружено соответствующих ячеек, свойству `isNullObject` объекта присваивается значение `true`.
- Оно вызывает `context.sync` *прежде*, чем протестировать свойство `isNullObject`. Это требование для всех методов и свойств `*OrNullObject`, так как всегда нужно загружать и синхронизировать свойство для его чтения. Тем не менее, не требуется *явным образом* нагружать свойство `isNullObject`. Он автоматически загружается с `context.sync` даже в том случае, если `load` не вызывается в объекте. Для получения дополнительных сведений см. [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).
- Этот код можно проверить, выбрав сначала диапазон, у которого нет ячеек формулы, и запустив его. Выберите диапазон, который содержит по крайней мере одну ячейку с формулой, и снова запустите его.

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

Для простоты во всех других примерах в этой статье используйте метод `getSpecialCells` вместо `getSpecialCellsOrNullObject`.

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>Сужение целевых ячеек с типом значений ячеек

Есть также необязательный второй параметр типа перечисления `Excel.SpecialCellValueType`, который в дальнейшем сужает ячейки до целевого объекта. Можно использовать его только в том случае, если передается значение "Formulas" или "Constants" для `getSpecialCells` или `getSpecialCellsOrNullObject`. Этот параметр указывает, что требуются только ячейки с определенными типами значений. Существует четыре основных типа: "Error", "Logical" (то же самое, что и boolean- логический), "Numbers" и "Text". (Перечисление имеет другие значения помимо этих четырех, которые рассматриваются ниже.) Ниже приведен пример. Вот что нужно знать об этом коде:

- Он будет выделять только ячейки, имеющие числовое значение литерала. Не выделяет ячейки, в которых содержится формула (даже в том случае, если результат является числом), логическое, текстовое значение, или ячейки с состоянием ошибки.
- Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

В некоторых случаях вам требуется работать с более чем одним типом значения ячейки, например, все ячейки с текстовыми значениями и все ячейки с логическими значениями ("Logical"). Перечисление `Excel.SpecialCellValueType` содержит значения, которые позволяют объединять типы. Например, "LogicalText" будет обрабатывать все логические и все текстовые ячейки. Можно использовать любые два или три из четырех основных типов. Имена этих значений перечисления, которые объединяют основные типы, всегда находятся в алфавитном порядке. Таким образом, для объединения ячеек со значениями ошибок, текстовыми и логическими значениями используйте "ErrorLogicalText", а не "LogicalErrorText" или "TextErrorLogical". Параметр по умолчанию "All" объединяет все четыре типа. В следующем примере выделены все ячейки с формулами, которые производят числовые или логические значения:

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> Параметр `Excel.SpecialCellValueType` можно использовать, только если параметр `Excel.SpecialCellType` – это "Formulas" или "Constants".

### <a name="get-rangeareas-within-rangeareas"></a>Получение объектов RangeAreas в RangeAreas

Тип `RangeAreas` также имеет методы `getSpecialCells` и `getSpecialCellsOrNullObject`, которые принимают те же два параметра. Эти методы возвращают все целевые ячейки из всех диапазонов в коллекции `RangeAreas.areas`. Существует одно небольшое отличие в поведении методов при вызове объекта `RangeAreas` вместо объекта `Range`: когда вы передаете "SameConditionalFormat" в качестве первого параметра, метод возвращает все ячейки, имеющие одинаковое условное форматирование, в качестве верхней крайней слева ячейки *в первом диапазоне в коллекции `RangeAreas.areas`*. То же касается и "SameDataValidation": при передаче к`Range.getSpecialCells` он возвращает все ячейки, которые имеют такое же правило проверки данных, в качестве верхней крайней слева ячейки *в диапазоне*. Но при передаче к `RangeAreas.getSpecialCells` она возвращает все ячейки, которые имеют такое же правило проверки данных, в качестве верхней крайней слева ячейки *в первый диапазон в коллекции `RangeAreas.areas`*.

## <a name="read-properties-of-rangeareas"></a>Чтение свойств RangeAreas

Чтение значения свойств `RangeAreas` требует осторожности, так как данное свойство может иметь разные значения для разных диапазонов в `RangeAreas`. Общее правило заключается в том, что если соответствующее значение *может* быть возвращено, оно будет возвращено. Например, в следующем коде RGB-код для розовой заливки (`#FFC0CB`) и `true` будет выполнять вход в консоль, так как оба диапазона в объекте `RangeAreas` имеют розовую заливку и оба являются целыми столбцами.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

Все усложняется, когда согласованность невозможна. Поведение свойств `RangeAreas` следует следующим тремя принципам:

- Логическое свойство объекта `RangeAreas` возвращает `false`, кроме случаев, когда свойство имеет значение true для всех диапазонов элементов.
- Свойства, не являющиеся логическими, за исключением свойства `address`, возвращают `null`, кроме тех случаев, когда соответствующее свойство для всех элементов диапазона обладает тем же значением.
- Свойство `address` возвращает строку с адресами диапазонов элементов, разделенными запятыми.

Например, следующий код создает `RangeAreas`, в котором только один диапазон — это целый столбец, и только один заполнен розовым цветом. Консоль покажет `null` для цвета заливки, `false` для свойства `isEntireRow` и "Sheet1! F3:F5 Sheet1! H:H" (при условии, что имя листа – это "Sheet1") для свойства`address`. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a>См. также

- [Основные понятия API JavaScript для Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [Объект Range (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- Объект[RangeAreas (JavaScript API для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (Эта ссылка может не работать, пока API находится в режиме предварительной версии. В качестве альтернативы см. [office.d.ts бета-версии](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)