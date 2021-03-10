---
title: Вызов API JavaScript Excel из настраиваемой функции
description: Узнайте, какие API JavaScript Excel можно вызвать из настраиваемой функции.
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613908"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>Вызов API JavaScript Excel из настраиваемой функции

Вызов API JavaScript Excel из пользовательских функций, чтобы получить данные о диапазоне и получить дополнительный контекст для вычислений. Вызов API JavaScript Excel с помощью настраиваемой функции может быть полезен, если:

- Перед вычислением настраиваемая функция должна получать сведения из Excel. Эти сведения могут включать свойства документов, форматы диапазона, пользовательские XML-части, имя книги или другую информацию, определенную в Excel.
- Настраиваемая функция будет устанавливать формат номера ячейки для возвращаемого значения после вычисления.

> [!IMPORTANT]
> Чтобы вызвать API JavaScript Excel из настраиваемой функции, необходимо использовать общее время запуска JavaScript. Дополнительные сведения см. в статье [Настройка надстройки Office для использования общей среды выполнения JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="code-sample"></a>Пример кода

Чтобы вызвать API JavaScript Excel из настраиваемой функции, сначала требуется контекст. Чтобы получить контекст, используйте объект [Excel.RequestContext.](/javascript/api/excel/excel.requestcontext) Затем используйте контекст для вызова API, необходимых в книге.

В следующем примере кода показано, как использовать для получения значения из `Excel.RequestContext` ячейки в книге. В этом примере параметр передается в метод `address` Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) и должен быть введен в качестве строки. Например, настраиваемая функция, вступив в пользовательский интерфейс Excel, должна следовать шаблону , где находится адрес ячейки, из которой можно получить `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` значение.

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>Ограничения вызовов API JavaScript Excel с помощью настраиваемой функции

Не вызывайте API JavaScript Excel из настраиваемой функции, которая меняет среду Excel. Это означает, что пользовательские функции не должны выполнять следующие функции:

- Вставка, удаление или форматирование ячеек в таблицу.
- Измените значение другой ячейки.
- Перемещение, переименование, удаление или добавление листов в книгу.
- Измените все параметры среды, такие как режим вычисления или представления экрана.
- Добавление имен в книгу.
- Установите свойства или выполните большинство методов.

Изменение Excel может привести к низкой производительности, выходу времени и бесконечным циклам. Настраиваемые вычисления функций не должны запускаться во время пересчета Excel, так как это приведет к непредсказуемым результатам.

Вместо этого внести изменения в Excel из контекста кнопки ленты или области задач.

## <a name="next-steps"></a>Дальнейшие действия

- [Основные концепции программирования с помощью API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>См. также

- [Совместное делиться данными и событиями между пользовательскими функциями Excel и учебником по области задач](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
