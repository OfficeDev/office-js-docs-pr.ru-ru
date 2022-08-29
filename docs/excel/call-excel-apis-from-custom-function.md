---
title: Вызов API JavaScript для Excel из пользовательской функции
description: Узнайте, какие API JavaScript для Excel можно вызывать из пользовательской функции.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa22cb007bb4803863c17e0f72876cc58c15b992
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423190"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>Вызов API JavaScript для Excel из пользовательской функции

Вызовите API JavaScript для Excel из пользовательских функций, чтобы получить данные диапазона и получить дополнительный контекст для вычислений. Вызов API JavaScript для Excel с помощью пользовательской функции может быть полезен, если:

- Перед вычислением пользовательская функция должна получить сведения из Excel. Эти сведения могут включать свойства документа, форматы диапазонов, пользовательские XML-части, имя книги или другие сведения, относящиеся к Excel.
- Пользовательская функция задает числовую формат ячейки для возвращаемых значений после вычисления.

> [!IMPORTANT]
> Чтобы вызвать API JavaScript для Excel из пользовательской функции, необходимо использовать общую [среду выполнения](../testing/runtimes.md#shared-runtime). [Дополнительные сведения см](../develop/configure-your-add-in-to-use-a-shared-runtime.md). в статье "Настройка надстройки Office для использования общей среды выполнения".

## <a name="code-sample"></a>Пример кода

Чтобы вызвать API JavaScript для Excel из пользовательской функции, сначала требуется контекст. Используйте объект [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) для получения контекста. Затем используйте контекст для вызова API, необходимых в книге.

В следующем примере кода показано `Excel.RequestContext` , как получить значение из ячейки книги. В этом примере `address` параметр передается в метод Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) и должен быть введен в виде строки. Например, пользовательская функция, введенная в пользовательский интерфейс Excel `=CONTOSO.GETRANGEVALUE("A1")`, должна соответствовать шаблону, `"A1"` где находится адрес ячейки, из которой извлекается значение.

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 const context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load("values");
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>Ограничения вызова API JavaScript для Excel с помощью пользовательской функции

Не вызывайте API JavaScript для Excel из пользовательской функции, которая изменяет среду Excel. Это означает, что пользовательские функции не должны выполнять следующие действия:

- Вставка, удаление или форматирование ячеек в электронной таблице.
- Измените значение другой ячейки.
- Перемещение, переименование, удаление или добавление листов в книгу.
- Измените любой из параметров среды, например режим вычисления или просмотр экрана.
- Добавьте имена в книгу.
- Задайте свойства или выполните большинство методов.

Изменение Excel может привести к снижению производительности, простоям и бесконечным циклам. Вычисления пользовательских функций не должны выполняться во время пересчета Excel, так как это приведет к непредсказуемым результатам.

Вместо этого внесите изменения в Excel из контекста кнопки ленты или области задач.

## <a name="next-steps"></a>Дальнейшие действия

- [Основные концепции программирования с помощью API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>См. также

- [Руководство по совместному доступу к данным и событиям между пользовательскими функциями Excel и областью задач](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Настройка надстройки Office для использования общей среды выполнения](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
