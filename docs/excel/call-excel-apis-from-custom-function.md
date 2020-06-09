---
title: Вызов API Microsoft Excel из настраиваемой функции
description: Узнайте, какие API Microsoft Excel вы можете вызывать из пользовательской функции.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: a25d3f151f648560ee24a3da3f689cb9767bd52a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609806"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a>Вызов API Microsoft Excel из настраиваемой функции

Вызовите API Office. js Excel из пользовательских функций, чтобы получить данные диапазона и получить дополнительный контекст для вычислений.

Вызов API Office. js с помощью настраиваемой функции может быть полезен в следующих случаях:

- Перед вычислением пользовательская функция должна получить сведения из Excel. Эти сведения могут включать в себя свойства документов, форматы диапазонов, пользовательские XML-части, имя книги или другие сведения, относящиеся к Excel.
- Настраиваемая функция будет задавать числовой формат ячейки для возвращаемых значений после вычисления.

## <a name="code-sample"></a>Пример кода

Для вызова API Office. js первым нужен контекст. Используйте `Excel.RequestContext` объект для получения контекста. Затем используйте контекст для вызова API, которые необходимы в книге.

В приведенном ниже примере кода показано, как получить диапазон значений из книги.

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a>Ограничения на вызов Office. js с помощью настраиваемой функции

Не вызывайте API Office. js из пользовательской функции, которая изменяет среду Excel. Это означает, что пользовательские функции не должны выполнять следующие действия:

- Вставка, удаление или форматирование ячеек в электронной таблице.
- Изменить значение другой ячейки.
- Перемещение, переименование, удаление и добавление листов в книгу.
- Измените любые параметры среды, такие как режим вычисления или экранные представления.
- Добавление имен в книгу.
- Задайте свойства или выполните большинство методов.

Изменение Excel может привести к ухудшению производительности, времени ожидания и бесконечному циклу. Пользовательские вычисления функций не должны выполняться во время пересчета Excel, так как это приведет к непредсказуемым результатам.

Вместо этого внесите изменения в Excel из контекста кнопки на ленте или области задач.

## <a name="next-steps"></a>Дальнейшие действия

- [Основные концепции программирования с помощью API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>См. также

- [Обмен данными и событиями между пользовательскими функциями и областью задач Excel](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
