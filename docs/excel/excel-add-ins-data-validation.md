---
title: Добавление проверки данных в диапазоны Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: e584d80403969ca15527afb69251832a96f3003e
ms.sourcegitcommit: 979e8d2b0fbc537f9c01527062767c026592aa62
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2018
ms.locfileid: "25852919"
---
# <a name="add-data-validation-to-excel-ranges"></a>Добавление проверки данных в диапазоны Excel

Библиотека JavaScript Excel предоставляет API, позволяющие вашей надстройке добавлять функцию автоматической проверки данных для таблиц, столбцов, строк и других диапазонов в книге.  Чтобы понять принципы и терминологию проверки данных, ознакомьтесь со следующими статьями о том, как пользователи добавляют проверку данных в пользовательском интерфейсе Excel:

- [Применение проверки данных к ячейкам](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Подробнее о проверке данных](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Описание и примеры проверки данных в Excel](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Программное управление проверкой данных

Свойство `Range.dataValidation`, которое получает объект [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation), является точкой входа для программного управления проверкой данных в Excel. Существует пять свойств объекта `DataValidation`:

- `rule` — определяет, какие данные для диапазона являются допустимыми. См. статью [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).
- `errorAlert` — указывает, появляется ли ошибка, если пользователь вводит недопустимые данные, и определяет текст, название и стиль оповещения, например **Informational** (информирование), **Warning** (предупреждение) и **Stop** (остановка).  См. статью [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).
- `prompt` — указывает, появляется ли подсказка, когда пользователь наводит указатель мыши на диапазон, и определяет текст подсказки. См. статью [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).
- `ignoreBlanks` — указывает, применяется ли правило проверки данных к пустым ячейкам в диапазоне. Значение по умолчанию: `true`.
- `type` — идентификация типа проверки "только для чтения", например WholeNumber, Date, TextLength и т. д. Это свойство устанавливается неявно при задании свойства `rule`.

> [!NOTE]
> Проверка данных, добавляемая программно, ведет себя так же, как проверка данных, добавляемая вручную.  В частности, обратите внимание на то, что проверка данных запускается только в том случае, если пользователь вводит значение в ячейку или копирует и вставляет ячейки из другого источника в книге и выбирает параметр вставки **Значения**. Если пользователь копирует ячейку и выполняет простую вставку в диапазон проверки данных, проверка не выполняется.

## <a name="creating-validation-rules"></a>Создание правил проверки

Чтобы добавить проверку данных в диапазон, ваш код должен установить свойство `rule` объекта `DataValidation` в `Range.dataValidation`. Это приводит к получению объекта [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule), который имеет семь дополнительных свойств. *Максимум одно свойство может присутствовать в любом объекте `DataValidationRule`.* Указываемое свойство определяет тип выполняемой проверки.

### <a name="basic-and-datetime-validation-rule-types"></a>Типы правил проверки Basic и DateTime

Первые три свойства `DataValidationRule` (т. е. типы правил проверки) в качестве своего значения принимают объект [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation).

- `wholeNumber` — требует целое число в дополнение к другим проверкам, указанным объектом `BasicDataValidation`.
- `decimal` — требует десятичное число в дополнение к другим проверкам, указанным объектом `BasicDataValidation`.
- `textLength` — применяет сведения проверки объекта `BasicDataValidation` к *длине* значения ячейки.

Ниже приведен пример создания правила проверки.  Обратите внимание на следующие особенности этого кода:

- `operator` — это бинарный оператор "GreaterThan". При использовании бинарного оператора значение, которое пользователь пытается ввести в ячейку, — это левый операнд, а значение, указанное в `formula1`, — это правый операнд. Поэтому согласно этому правилу только целые числа больше 0 являются допустимыми. 
- `formula1` — это жестко заданное число. Если во время кодирования вы не знаете, какое значение должно быть задано, можно также использовать формулу Excel (в виде строки) для значения. Например, "= A3" и "= SUM(A4,B5)" могут также быть значениями `formula1`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            wholeNumber: {
                formula1: 0,
                operator: "GreaterThan"
            }
        };

    return context.sync();
})
```

Перечень других бинарных операторов см. в статье [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation). 

Существует также два тернарных оператора: "Between" и "NotBetween". Для их использования необходимо указать необязательное свойство `formula2`.  Значения `formula1` и `formula2` — это ограничивающие операнды. Значение, которое пользователь пытается ввести в ячейку, — это третий (вычисленный) операнд. Ниже приведены примеры использования оператора "Between":

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            decimal: {
                formula1: 0,
                formula2: 100,
                operator: "Between"
            }
        };

    return context.sync();
})
```

Следующие два свойства правила в качестве своего значения принимают объект [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation).

- `date`
- `time`

Объект `DateTimeDataValidation` структурирован так же, как и `BasicDataValidation`: он имеет свойства `formula1`, `formula2` и `operator` и используется аналогичным образом. Различие состоит в том, что в свойствах формулы нельзя использовать число, но можно ввести строку [даты и времени ISO 8606](https://www.iso.org/iso-8601-date-and-time-format.html) (или формулу Excel). Ниже приведен пример, в котором определяются допустимые значения для дат в первую неделю апреля 2018 года. 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            date: {
                formula1: "2018-04-01",
                formula2: "2018-04-08",
                operator: "Between"
            }
        };

    return context.sync();
})
```

### <a name="list-validation-rule-type"></a>Тип правила проверки для списка

Используйте свойство `list` в объекте `DataValidationRule`, чтобы указать, что единственными допустимыми значениями являются значения из конечного списка. Ниже приведен пример. Обратите внимание на следующие особенности этого кода:

- Предполагается, что существует лист с именем "Имена", а значения в диапазоне "A1:A3" являются именами.
- Свойство `source` определяет список допустимых значений. Строковый аргумент ссылается на диапазон с именами. Можно также назначить разделенный запятыми список, например "Регина, Сергей, Анна". 
- Свойство `inCellDropDown` указывает, будет ли раскрывающийся элемент управления отображаться в ячейке, когда пользователь выбирает ее. Если свойству присвоено значение `true`, то раскрывающийся список отображается со списком значений из `source`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: "=Names!$A$1:$A$3"
        }
    };

    return context.sync();
})
```

### <a name="custom-validation-rule-type"></a>Настраиваемый тип правила проверки

Используйте свойство `custom` в объекте `DataValidationRule`, чтобы задать настраиваемую формулу проверки. Ниже приведен пример. Обратите внимание на следующие особенности этого кода:

- Предполагается, что на листе расположена таблица с двумя столбцами A и B: **Имя спортсмена** и **Комментарии**.
- Чтобы исключить многословие в столбце **Комментарии**, данные, содержащие имя спортсмена, определяются недопустимыми.
- `SEARCH(A2,B2)` возвращает стартовую позицию строки в ячейке A2 в строку в ячейке B2. Если A2 не находится в ячейке B2, не возвращается числовое значение. `ISNUMBER()` возвращает логическое значение. Поэтому свойство `formula` указывает, что допустимые данные для столбца **Комментарии** — это данные, которые не содержат строку в столбце **Имя спортсмена**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var commentsRange = sheet.tables.getItem("AthletesTable").columns.getItem("Comments").getDataBodyRange();

    commentsRange.dataValidation.rule = {
            custom: {
                formula: "=NOT(ISNUMBER(SEARCH(A2,B2)))"
            }
        };

    return context.sync();
})
```

## <a name="create-validation-error-alerts"></a>Создание оповещений об ошибках проверки

Вы можете создать настраиваемое оповещение об ошибке, которое отображается, если пользователь пытается ввести недопустимые данные в ячейке. Ниже приведен простой пример. Обратите внимание на следующие особенности этого кода:

- Свойство `style` определяет, получает ли пользователь информационное уведомление, предупреждение или оповещение "stop". Только `Stop` действительно не позволяет пользователю добавлять недопустимые данные.  Всплывающее окно для `Warning` и `Information` содержит параметры, позволяющие пользователю в любом случае ввести недопустимые данные.
- Свойству `showAlert` по умолчанию присвоено значение `true`.  Это означает, что узел Excel будет выдавать всплывающее окно с универсальным оповещением (типа `Stop`), если вы не создадите настраиваемое оповещение, которое либо устанавливает для `showAlert` значение `false`, либо задает настраиваемое сообщение, заголовок и стиль.  Этот код задает настраиваемое сообщение и заголовок.


```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.errorAlert = {
            message: "Sorry, only positive whole numbers are allowed",
            showAlert: true, // default is 'true'
            style: "Stop", // other possible values: Warning, Information
            title: "Negative or Decimal Number Entered"
        };
    
    // Set range.dataValidation.rule and optionally .prompt here.

    return context.sync();
})
```

Дополнительные сведения см. в статье [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).

## <a name="create-validation-prompts"></a>Создание запросов проверки

Вы можете создать пояснительную подсказку, которая появляется, когда пользователь наводит указатель мыши на ячейку, к которой была применена проверка данных, или выбирает ее. Ниже приведен пример.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.prompt = {
            message: "Please enter a positive whole number.",
            showPrompt: true, // default is 'false'
            title: "Positive Whole Numbers Only."
        };
    
    // Set range.dataValidation.rule and optionally .errorAlert here.

    return context.sync();
})
```

Дополнительные сведения см. в статье [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).

## <a name="remove-data-validation-from-a-range"></a>Удаление проверки данных из диапазона

Чтобы удалить проверку данных из диапазона, вызовите метод [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--).

```js
myrange.dataValidation.clear()
```

Необязательно, чтобы очищаемый диапазон был тем же диапазоном, к которому вы применили проверку данных. Если это не один и тот же диапазон, удаляются только перекрывающиеся ячейки двух диапазонов (при их наличии). 

> [!NOTE]
> Удаление проверки данных из диапазона также распространяется на любую проверку данных, которую пользователь добавил вручную в диапазон.

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Объект DataValidation (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [Объект Range (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
