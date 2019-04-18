---
title: Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)
description: ''
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: aacbe930e2cf3da4d10b61bfe8f34efe1094c113
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914243"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a>Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)

Эта статья основана на сведениях из статьи [Работа с диапазонами с использованием API JavaScript для Excel (основные задачи)](excel-add-ins-ranges.md) с предоставлением примеров кода, демонстрирующих способы выполнения более сложных задач с диапазонами с использованием API JavaScript для Excel. Полный список свойств и методов, поддерживаемых объектом **Range**, см. в статье [Объект Range (API JavaScript для Excel)](/javascript/api/excel/excel.range).

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a>Работа с датами с использованием подключаемого модуля Moment-MSDate

[Библиотека JavaScript Moment](https://momentjs.com/) предоставляет удобный способ использования дат и меток времени. [Подключаемый модуль Moment-MSDate](https://www.npmjs.com/package/moment-msdate) преобразует формат моментов времени в предпочитаемый для Excel. Это тот же формат, который возвращает [функция ТДАТА](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46).

В приведенном ниже коде показано, как установить для диапазона в **B4** метку момента времени.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

Это похоже на способ получения даты из ячейки и ее преобразования в формат момента времени или другой формат, как показано в приведенном ниже коде:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

Вашей надстройке потребуется отформатировать диапазоны, чтобы отобразить даты в более понятной для человека форме. В примере `"[$-409]m/d/yy h:mm AM/PM;@"` время отобразится как "12/3/18 3:57 PM". Дополнительные сведения о форматах чисел даты и времени см. в разделе "Рекомендации по форматам даты и времени" статьи [Рекомендации по настройке числовых форматов](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).

## <a name="work-with-multiple-ranges-simultaneously-preview"></a>Работа с несколькими диапазонами одновременно (предварительная версия)

> [!NOTE]
> `RangeAreas` Объект в настоящее время доступен только в общедоступной предварительной версии. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

Объект `RangeAreas` позволяет вашей надстройке выполнять операции над несколькими диапазонами одновременно. Эти диапазоны могут быть смежными, но это необязательно. Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).

## <a name="find-special-cells-within-a-range-preview"></a>Поиск специальных ячеек в диапазоне (предварительная версия)

> [!NOTE]
> Методы `getSpecialCells` и `getSpecialCellsOrNullObject` в настоящее время доступны только в общедоступной предварительной версии. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` находят диапазоны с учетом характеристик ячеек и типов значений ячеек. Оба этих метода возвращают объекты `RangeAreas`. Подписи методов из файла типов данных TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

В приведенном ниже примере используется метод `getSpecialCells`, чтобы найти все ячейки с формулами. Вот что нужно знать об этом коде:

- Он ограничивает часть листа, в которой требуется выполнять поиск, путем вызова сначала метода `Worksheet.getUsedRange`, а затем метода `getSpecialCells` только для этого диапазона.
- Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами окрашены розовым цветом даже в том случае, если они не являются смежными.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

Если в диапазоне нет ячеек с целевыми характеристиками, метод `getSpecialCells` выдает ошибку **ItemNotFound**. Это приведет к переадресации потока управления к блоку `catch`, если таковой существует. Если блок `catch` отсутствует, ошибка останавливает исполнение функции.

Если ожидается, что всегда должны существовать ячейки с целевыми характеристиками, скорее всего вы захотите, чтобы код выдавал ошибку при их отсутствии. Если отсутствие соответствующих ячеек является допустимым сценарием, ваш код должен проверить наличие такой возможности и корректно выполнить действие без выдачи ошибки. Добиться такого поведения можно с помощью метода `getSpecialCellsOrNullObject` и возвращаемого им свойства `isNullObject`. Этот шаблон используется в приведенном ниже примере. Вот что нужно знать об этом коде:

- Метод `getSpecialCellsOrNullObject` всегда возвращает прокси-объект, поэтому он не может иметь значение `null` в обычном смысле JavaScript. Но если соответствующие ячейки не обнаружены, свойству `isNullObject` объекта присваивается значение `true`.
- Он вызывает `context.sync` *перед* тестированием свойства `isNullObject`. Это требование для всех методов и свойств `*OrNullObject`, так как всегда нужно загружать и синхронизировать свойство, чтобы его прочесть. Однако необязательно *явно* загружать свойство `isNullObject`. Оно автоматически загружается с помощью `context.sync`, даже если `load` не вызывается для объекта. Дополнительные сведения см. в разделе [\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).
- Этот код можно проверить, выбрав сначала диапазон без ячеек с формулами и запустив его. Затем следует выбрать диапазон, содержащий по крайней мере одну ячейку с формулой, и снова запустить его.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
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

Для удобства во всех других примерах в этой статье используйте метод `getSpecialCells` вместо `getSpecialCellsOrNullObject`.

### <a name="narrow-the-target-cells-with-cell-value-types"></a>Ограничение целевых ячеек с помощью типа значений ячеек

Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` принимают необязательный второй параметр, используемый для дополнительного ограничения целевых ячеек. Этот второй параметр `Excel.SpecialCellValueType` используется для указания того, что требуются только ячейки, содержащие определенные типы значений.

> [!NOTE]
> Параметр `Excel.SpecialCellValueType` можно использовать, только если для параметра `Excel.SpecialCellType` задано значение `Excel.SpecialCellType.formulas` или `Excel.SpecialCellType.constants`.

#### <a name="test-for-a-single-cell-value-type"></a>Тестирование для ячеек с одним типом значений

Для перечисления `Excel.SpecialCellValueType` существует четыре основных типа (в дополнение к другим объединенным значениям, описанным ниже в этом разделе):

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical` (означает логическое значение)
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

В приведенном ниже примере выполняется поиск специальных ячеек, являющихся числовыми константами, и их окрашивание в розовый цвет. Вот что нужно знать об этом коде:

- Он выделяет только ячейки с числовым значением литерала. Он не выделяет ячейки с формулой (даже если результат является числом), логическим значением, текстовым значением или ячейки с состоянием ошибки.
- Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

#### <a name="test-for-multiple-cell-value-types"></a>Тестирование для ячеек с несколькими типами значений

Иногда требуется работать с ячейками, имеющими несколько типов значений, например со всеми ячейками с текстовыми значениями и всеми ячейками с логическими значениями (`Excel.SpecialCellValueType.logical`). Для перечисления `Excel.SpecialCellValueType` существуют значения с объединенными типами. Например, `Excel.SpecialCellValueType.logicalText` обрабатывает все ячейки с логическими и текстовыми значениями. `Excel.SpecialCellValueType.all` является значением по умолчанию, которое не ограничивает возвращаемые типы значений ячеек. В приведенном ниже примере окрашены все ячейки с формулами, которые производят числовое или логическое значение.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="copy-and-paste-preview"></a>Копирование и вставка (предварительная версия)

> [!NOTE]
> Функция `Range.copyFrom` в настоящее время доступна только в общедоступной предварительной версии. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

Функция `copyFrom` диапазона реплицирует поведение копирования и вставки пользовательского интерфейса Excel. Диапазон объекта, который вызывается `copyFrom`, является назначением.
Источник для копирования передается как диапазон или адрес строки, представляющий диапазон.
В следующем примере кода копируются данные из **A1:E1** в диапазон, начиная с **G1** (который заканчивается вставкой в **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

У функции `Range.copyFrom` есть три необязательных параметра.

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` указывает, какие данные копируются из источника в назначение.

- `Excel.RangeCopyType.formulas` переносит формулы в ячейках источника и сохраняет относительное положение диапазонов этих формул. Все записи, не являющиеся формулами, копируются в исходном виде.
- `Excel.RangeCopyType.values` копирует значения данных, а в случае формул — результат формулы.
- `Excel.RangeCopyType.formats` копирует форматирование диапазона, включая шрифт, цвет и другие параметры форматирования, но без значений.
- `Excel.RangeCopyType.all` (вариант по умолчанию) копирует данные и форматирование, сохраняя формулы ячеек при их обнаружении.

`skipBlanks` устанавливает, будут ли копироваться пустые ячейки в назначение. Если значение равно true, `copyFrom` пропускает пустые ячейки в диапазоне источника.
Пропущенные ячейки не перезапишут существующие данные в соответствующих им ячейках конечного диапазона. Значение по умолчанию: false.

`transpose` определяет, переставляются ли данные в исходное расположение, то есть переключаются ли строки и столбцы.
Переставленный диапазон переключается на главной диагонали, поэтому строки **1**, **2** и **3** становятся столбцами **A**, **B** и **C**.

В приведенном ниже примере кода и изображениях демонстрируется это поведение в простом сценарии.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*Прежде чем предыдущая функция была запущена.*

![Данные в Excel перед запуском метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-before.png)

*После запуска предыдущей функции.*

![Данные в Excel после запуска метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates-preview"></a>Удаление дубликатов (предварительная версия)

> [!NOTE]
> Функция `removeDuplicates` объекта Range в настоящее время доступна только в общедоступной предварительной версии. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

Функция `removeDuplicates` объекта Range удаляет строки с повторяющимися записями в указанных столбцах. Функция проверяет каждую строку в диапазоне от индекса с наименьшим значением до индекса с наибольшим значением (сверху вниз). Строка удаляется, если значение в ее указанном столбце или столбцах уже встречалось в диапазоне. Строки в диапазоне под удаленной строкой сдвигаются вверх. Функция `removeDuplicates` не влияет на положение ячеек вне диапазона.

Функция `removeDuplicates` использует параметр `number[]`, представляющий индексы столбцов, которые проверяются на наличие дубликатов. Этот массив отсчитывается от нуля относительно диапазона, а не листа. Функция также использует логический параметр, который определяет, является ли первая строка заголовком. При значении **true** верхняя строка игнорируется при поиске дубликатов. Функция `removeDuplicates` возвращает объект `RemoveDuplicatesResult`, указывающий количество удаленных строк и количество оставшихся уникальных строк.

При использовании функции `removeDuplicates` диапазона, учитывайте следующее:

- Функция `removeDuplicates` рассматривает значения ячеек, а не результаты функций. Если две разные функции вычисляют одинаковый результат, значения ячеек не считаются повторяющимися.
- Пустые ячейки не игнорируются функцией `removeDuplicates`. Значение пустой ячейки обрабатывается как любое другое значение. Это означает, что пустые строки, содержащиеся в диапазоне, будут включены в объект `RemoveDuplicatesResult`.

В приведенном ниже примере показано удаление записей с повторяющимися значениями в первом столбце.

```js
Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*Прежде чем предыдущая функция была запущена.*

![Данные в Excel перед запуском метода удаления дубликатов](../images/excel-ranges-remove-duplicates-before.png)

*После запуска предыдущей функции.*

![Данные в Excel после запуска метода удаления дубликатов](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>См. также

- [Работа с диапазонами с использованием API JavaScript для Excel](excel-add-ins-ranges.md)
- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
