---
title: Excel Набор требований для API javaScript только для интернета
description: Сведения о наборе требований ExcelApiOnline.
ms.date: 10/13/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ae014930d3ec11d52b3904ee1205b670f8d3790f
ms.sourcegitcommit: 3b187769e86530334ca83cfdb03c1ecfac2ad9a8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/15/2021
ms.locfileid: "60367329"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel Набор требований для API javaScript только для интернета

Набор требований — это специальный набор требований, который включает функции, доступные только для `ExcelApiOnline` Excel в Интернете. API в этом наборе требований считаются производственными API (не подверженными незадокументированные поведенческие или структурные изменения) для Excel в Интернете приложения. `ExcelApiOnline`API считаются API предварительного просмотра для других платформ (Windows, Mac, iOS) и не могут поддерживаться ни одной из этих платформ.

Когда API в наборе требований поддерживаются на всех платформах, они будут добавлены в следующий выпущенный `ExcelApiOnline` набор требований ( `ExcelApi 1.[NEXT]` ). После того как это новое требование станет общедоступным, эти API будут удалены из `ExcelApiOnline` . Думайте об этом как об аналогичном процессе продвижения по службе aPI, перемещаемом с предварительного просмотра на выпуск.

> [!IMPORTANT]
> `ExcelApiOnline` — это суперсет последнего набора требований с номерами.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` это единственная версия API только для интернета. Это потому, Excel в Интернете всегда будет иметь одну версию, доступную пользователям, которая является последней версией.

В следующей таблице приводится краткий сводка [API,](#api-list) а в следующей таблице списка API приводится подробный список текущих `ExcelApiOnline` API.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Связанные книги | Управление связями между книгами, включая поддержку обновления и разрыва ссылок на книги. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Именуемые представления листа | Предоставляет программный контроль представлений таблицы для каждого пользователя. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |

## <a name="recommended-usage"></a>Рекомендуемое использование

Так как API поддерживаются только Excel в Интернете, надстройка должна проверить, поддерживается ли набор требований перед вызовом `ExcelApiOnline` этих API. Это позволяет избежать вызова API только для интернета на другой платформе.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

После того как API находится в наборе требований к платформе, необходимо удалить или изменить `isSetSupported` проверку. Это позволит включить функцию надстройки на других платформах. При внесении этого изменения обязательно проверьте эту функцию на этих платформах.

> [!IMPORTANT]
> Манифест не может `ExcelApiOnline 1.1` указываться как требование активации. Это значение не является допустимым для использования в [элементе Set.](../manifest/set.md)

## <a name="api-list"></a>Список API

В следующей таблице перечислены Excel API JavaScript, включенные в набор `ExcelApiOnline` требований. Полный список всех API Excel JavaScript (включая API и ранее выпущенные API), см. Excel `ExcelApiOnline` [API JavaScript.](/javascript/api/excel?view=excel-js-online&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|Очищает критерии фильтрации столбцов автофайлов.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|Делает запрос на разрыв ссылок, указывающих на связанную книгу.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|Исходный URL-адрес, указывающий на связанную книгу.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|Делает запрос на обновление данных, извлеченных из связанной книги.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|Нарушает все ссылки на связанные книги.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|Получает сведения о связанной книге по URL-адресу.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|Получает сведения о связанной книге по URL-адресу.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|Делает запрос на обновление всех ссылок на книги.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|Представляет режим обновления ссылок на книги.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate__)|Активирует это представление листа.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete__)|Удаляет представление листа из листа.|
||[дубликат (имя?: строка)](/javascript/api/excel/excel.namedsheetview#duplicate_name_)|Создает копию этого представления листа.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Получает или задает имя представления листа.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add_name_)|Создает новое представление листа с заданным именем.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#enterTemporary__)|Создает и активирует новое временное представление листа.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit__)|Выходит из действующего представления листа.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getActive__)|Получает в настоящее время активное представление листа.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getCount__)|Получает количество просмотров листов в этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItem_key_)|Получает представление листа с его именем.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getItemAt_index_)|Получает представление листа по индексу в коллекции.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|Удаление нескольких строк из таблицы.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|Удаление указанного количества строк из таблицы, начиная с указанного индекса.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Возвращает коллекцию связанных книг.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedSheetViews)|Возвращает коллекцию представлений листов, присутствующих в листе.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#onNameChanged)|Возникает при смене имени таблицы.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#onVisibilityChanged)|Возникает при смене видимости таблицы.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#onMoved)|Возникает при перемещении таблицы пользователем в книге.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#onNameChanged)|Возникает при смене имени таблицы в коллекции таблиц.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#onVisibilityChanged)|Возникает при смене видимости таблицы в коллекции таблиц.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#positionAfter)|Получает новую позицию таблицы после перемещения.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#positionBefore)|Получает предыдущую позицию таблицы перед перемещением.|
||[источник](/javascript/api/excel/excel.worksheetmovedeventargs#source)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#worksheetId)|Получает ID перемещенного таблицы.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameAfter)|Получает новое имя таблицы после изменения имени.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#nameBefore)|Получает предыдущее имя таблицы до изменения имени.|
||[источник](/javascript/api/excel/excel.worksheetnamechangedeventargs#source)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#worksheetId)|Получает ID таблицы с новым именем.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[источник](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#source)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#type)|Получает тип события.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityAfter)|Получает новый параметр видимости таблицы после изменения видимости.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#visibilityBefore)|Получает предыдущий параметр видимости таблицы перед изменением видимости.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#worksheetId)|Получает ID таблицы, видимость которой изменилась.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Предварительные версии API JavaScript для Excel](excel-preview-apis.md)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
