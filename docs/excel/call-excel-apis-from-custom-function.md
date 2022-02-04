---
title: Вызов Excel API JavaScript из настраиваемой функции
description: 'Узнайте, какие Excel API JavaScript можно вызвать из настраиваемой функции.'
ms.date: 08/30/2021
ms.localizationpriority: medium
---

# <a name="call-excel-javascript-apis-from-a-custom-function"></a>Вызов Excel API JavaScript из настраиваемой функции

Вызов Excel API JavaScript из пользовательских функций, чтобы получить данные о диапазоне и получить дополнительный контекст для вычислений. Вызов Excel API JavaScript с помощью настраиваемой функции может быть полезен, если:

- Настраиваемая функция должна получать сведения из Excel до вычисления. Эти сведения могут включать свойства документов, форматы диапазонов, пользовательские XML-части, имя книги или другую Excel сведения.
- Настраиваемая функция будет устанавливать формат номера ячейки для возвращаемого значения после вычисления.

> [!IMPORTANT]
> Чтобы вызвать Excel API JavaScript из настраиваемой функции, необходимо использовать совместное время запуска JavaScript. Дополнительные сведения см. в статье [Настройка надстройки Office для использования общей среды выполнения JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="code-sample"></a>Пример кода

Чтобы вызвать Excel API JavaScript из настраиваемой функции, сначала требуется контекст. Используйте [Excel. Запрос объектаContext](/javascript/api/excel/excel.requestcontext) для получения контекста. Затем используйте контекст для вызова API, необходимых в книге.

В следующем примере кода показано, `Excel.RequestContext` как использовать для получения значения из ячейки в книге. В этом примере `address` параметр передается в метод Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) и должен быть введен в качестве строки. Например, настраиваемая функция, вступив Excel пользовательского `=CONTOSO.GETRANGEVALUE("A1")`интерфейса, должна следовать шаблону , `"A1"` где находится адрес ячейки, из которой можно получить значение.

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
 range.load("values");
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>Ограничения вызовов Excel API JavaScript с помощью настраиваемой функции

Не вызывайте Excel API JavaScript из настраиваемой функции, которая меняет среду Excel. Это означает, что пользовательские функции не должны выполнять следующие функции:

- Вставка, удаление или форматирование ячеек в таблицу.
- Измените значение другой ячейки.
- Перемещение, переименование, удаление или добавление листов в книгу.
- Измените все параметры среды, такие как режим вычисления или представления экрана.
- Добавление имен в книгу.
- Установите свойства или выполните большинство методов.

Изменение Excel может привести к низкой производительности, выходу времени и бесконечным циклам. Настраиваемые вычисления функций не должны запускаться во время Excel перерасчета, так как это приведет к непредсказуемым результатам.

Вместо этого внести изменения Excel из контекста кнопки ленты или области задач.

## <a name="next-steps"></a>Дальнейшие действия

- [Основные концепции программирования с помощью API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>См. также

- [Обмениваться данными и событиями между Excel пользовательскими функциями и учебником по области задач](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
