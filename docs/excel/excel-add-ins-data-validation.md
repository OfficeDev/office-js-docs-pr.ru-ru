---
title: Добавление проверки данных в диапазоны Excel
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: 3d6a901e2f8296806cff470340b40f4d77e79e34
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703947"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a>Добавление проверки данных в диапазоны Excel (предварительная версия)

> [!NOTE]
> Пока API проверки данных являются предварительной версией, для их использования вы должны загрузить бета-версию библиотеки JavaScript для Office. URL-адрес: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Если вы используете TypeScript или редактор кода использует файл определения типа TypeScript для IntelliSense, воспользуйтесь https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

> [!NOTE]
> Поскольку API проверки достоверности данных находятся в режиме предварительного просмотра, в этой статье ссылки на API работать не будут. Тем временем вы можете использовать [черновой вариант справки по API Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel).

Библиотека JavaScript Excel предоставляет API, позволяющие вашей надстройке добавлять автоматическую проверку данных в таблицы, столбцы, строки и другие диапазоны в книге. Для ознакомления с понятиями и терминологией проверки данных смотрите следующие статьи о том, как пользователи добавляют проверку данных через интерфейс Excel:

- [Применение проверки данных к ячейкам](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [Подробнее о проверке данных](https://support.office.com/en-us/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Описание и примеры проверки данных в Excel](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>Программный элемент управления проверкой данных

Свойство `Range.dataValidation`, которое принимает объект [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation), является точкой входа для программного управления проверкой данных в Excel. Существует пять свойств объекта `DataValidation`:

- `rule` — Определяет, какие данные для диапазона являются достоверными. См. [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).
- `errorAlert` — Указывает, появляется ли ошибка, если пользователь вводит неверные данные, и определяет текст, название и стиль оповещения, например, **Informational** (информирование), **Warning** (предупреждение), а также **Stop** (остановка). См. [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).
- `prompt` — Указывает, появляется ли подсказка, когда пользователь наводит курсор мыши на диапазон, и определяет текст подсказки. См. [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).
- `ignoreBlanks` — Указывает, применяется ли правило проверки данных к пустым ячейкам в диапазоне. Значение по умолчанию: `true`.
- `type` — Идентификация типа проверки "только для чтения", например, WholeNumber, Date, TextLength и т. д. Это свойство устанавливается неявно при установке `rule`.

> [!NOTE]
> Добавленная программно проверка данных ведет себя так же, как и добавленная вручную. Например, обратите внимание, что проверка данных запускается только в том случае, если пользователь непосредственно вводит значение в ячейку или копирует и вставляет ячейку из другого места в книге с параметром вставки **Значения**. Если пользователь копирует ячейку и выполняет простую вставку в диапазон с проверкой данных, проверка не запускается.

### <a name="creating-validation-rules"></a>Создание правил проверки

Чтобы добавить проверку данных в диапазон, в код нужно добавить свойство `rule` объекта `DataValidation` в `Range.dataValidation`. При этом используется объект [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule), который имеет семь необязательных свойств. *В любом объекте `DataValidationRule` может использоваться не более одного из этих свойств.* Включенное свойство определяет тип проверки.

#### <a name="basic-and-datetime-validation-rule-types"></a>Типы правил проверки Basic и DateTime

Первые три свойства `DataValidationRule` (т. е. типа правил проверки) в качестве своего значения принимают объект [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation).

- `wholeNumber` — Требует целое число в дополнение к другим проверкам, определенным в объекте `BasicDataValidation`.
- `decimal` — Требует десятичное число в дополнение к другим условиям проверки, определенным в объекте `BasicDataValidation`.
- `textLength` — Применяет сведения проверки объекта `BasicDataValidation` к *длине* значения ячейки.

Вот пример создания правила проверки. Обратите внимание на следующие особенности этого кода:

- — это бинарный оператор "GreaterThan".`operator` Всякий раз, когда вы используете бинарный оператор, значение, которое пользователь пытается ввести в ячейку, — это левый операнд, а значение, указанное в `formula1` — правый операнд. Таким образом, это правило устанавливает, что действительны только целые числа, которые больше 0. 
- — жестко заданное число.`formula1` Если во время написания кода вы не знаете, каким должно быть значение, для него можно использовать формулу Excel (как строку). Например, "= A3" и "= SUM (A4, B5)" также могут быть значениями `formula1`.

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

Есть также два тернарных оператора: "Between" и "NotBetween". Чтобы их использовать, необходимо указать необязательное свойство `formula2`. Значения `formula1` и `formula2` — ограничивающие операнды. Значение, которое пользователь вводит в ячейку, является третьим (оцениваемым) операндом. Ниже приведен пример использования оператора "Between":

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

Объект `DateTimeDataValidation` структурирован аналогично `BasicDataValidation`: он содержит свойства `formula1`, `formula2` и `operator` и используется таким же образом. Разница состоит в том, что число в свойствах формулы использовать нельзя, но можно ввести строку [с датой и временем по ISO 8606](https://www.iso.org/iso-8601-date-and-time-format.html) (или формулу Excel). Ниже приведен пример, в котором определены допустимые значения дат для первой недели апреля 2018 года. 

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

#### <a name="list-validation-rule-type"></a>Тип правила проверки List

Используйте свойство `list` для объекта `DataValidationRule` для указания того, что единственно допустимыми являются значения из ограниченного списка. Ниже приведен пример. Обратите внимание на следующие особенности этого кода:

- Предполагается, что существует лист с именем "Names", и значения в диапазоне "A1: A3" являются именами.
- Свойство `source`задает список допустимых значений. Ему присвоен диапазон с именами. Также можно назначить список с разделителями-запятыми, например, "Сью, Рики, Лиз". 
- Свойство `inCellDropDown` определяет, будет ли выпадающий элемент управления появляться в ячейке, когда пользователь ее выберет. Если установлено значение `true`, появится выпадающий список значений из `source`.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: nameSourceRange
        }
    };

    return context.sync();
})
```

#### <a name="custom-validation-rule-type"></a>Тип правила проверки Custom

Используйте свойство `custom` для объекта `DataValidationRule`, чтобы указать настраиваемую формулу проверки. Ниже приведен пример. Обратите внимание на следующие особенности этого кода:

- Предполагается, что на листе расположена таблица с двумя столбцами, A и B: **Athlete Name** (имя спортсмена) и **Comments** .
- Для исключения многословия в столбце **Комментарии** код определяет недопустимыми данные, которые содержат имя спортсмена.
- `SEARCH(A2,B2)` возвращает начальную позицию строки в A2 в строке в B2. Если A2 не содержится в B2, число не возвращается. `ISNUMBER()` возвращает логическое значение. Итак, свойство `formula` говорит, что данные в столбце**Comment** действительны, если в них не включена строка из столбца **Имя атлета**.

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

### <a name="create-validation-error-alerts"></a>Создание предупреждений об ошибках проверки

Вы можете создать настраиваемое предупреждение об ошибке, которое появляется, когда пользователь пытается ввести недопустимые данные в ячейку. Ниже приведен простой пример. Обратите внимание на следующие особенности этого кода:

- Свойство `style` определяет, какое сообщение получит пользователь: alert (оповещение), warning (предупреждение) или "stop" (стоп-оповещение). Только `Stop` действительно предотвращает добавление пользователем недопустимых данных. Всплывающее окна `Warning` и `Information` обладают параметрами, которые позволяют пользователю все равно ввести недопустимые данные.
- Свойство `showAlert` по умолчанию имеет значение `true`. Это означает, что в ведущем приложении Excel появится общее оповещение (типа `Stop`), если вы не создали настраиваемое оповещение, которое либо устанавливает `showAlert` значение `false`, либо устанавливает настраиваемое сообщение, заголовок и стиль. Этот код устанавливает настраиваемое сообщение и заголовок.


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

### <a name="create-validation-prompts"></a>Создание запросов проверки

Вы можете создать подсказку, которая появляется, когда пользователь наводит курсор мыши на ячейку, к которой применяется проверка данные, или выбирает ее. Ниже приведен пример.

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

### <a name="remove-data-validation-from-a-range"></a>Удаление проверки данных из диапазона

Чтобы удалить проверку данных из диапазона, вызовите метод [Range.dataValidation.clear ()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear).

```js
myrange.dataValidation.clear()
```

Не обязательно, чтобы диапазон, который вы очищаете, полностью совпадал с диапазоном, для которого вы добавили проверку данных. Если они не совпадают, очищаются только из двух диапазонов, которые совпадают. 

> [!NOTE]
> Очистка проверки данных из диапазона также распространяется на любую проверку данных, которую пользователь добавил вручную в диапазон.

## <a name="see-also"></a>См. также

- [Основные понятия API JavaScript для Excel](excel-add-ins-core-concepts.md)
- [Объект DataValidation (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [Объект Range (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
