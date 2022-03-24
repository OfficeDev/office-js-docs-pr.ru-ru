---
title: Excel API JavaScript только для интернета
description: Сведения о наборе требований ExcelApiOnline.
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f3ec510e889ecfe565767352c59cd349e0701830
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746605"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel API JavaScript только для интернета

Набор `ExcelApiOnline` требований — это специальный набор требований, который включает функции, доступные только для Excel в Интернете. API в этом наборе требований считаются производственными API (не подверженными незадокументированные поведенческие или структурные изменения) для Excel в Интернете приложения. `ExcelApiOnline`API считаются "предварительными" API для других платформ (Windows, Mac, iOS) и не могут поддерживаться ни одной из этих платформ.

Когда API в наборе `ExcelApiOnline` требований поддерживаются на всех платформах, они будут добавлены в следующий выпущенный набор требований (`ExcelApi 1.[NEXT]`). После того как это новое требование станет общедоступным, эти API будут удалены из `ExcelApiOnline`. Думайте об этом как об аналогичном процессе продвижения по службе aPI, перемещаемом с предварительного просмотра на выпуск.

> [!IMPORTANT]
> `ExcelApiOnline` — это суперсет последнего набора требований с номерами.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` это единственная версия API только для интернета. Это потому, Excel в Интернете всегда будет иметь одну версию, доступную пользователям, которая является последней версией.

В следующей таблице приводится краткий сводка API, а в следующей таблице списка [API](#api-list) приводится подробный список текущих `ExcelApiOnline` API.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Связанные книги | Управление связями между книгами, включая поддержку обновления и разрыва ссылок на книги. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Именуемые представления листа | Предоставляет программный контроль представлений таблицы для каждого пользователя. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview), [NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection) |
| События перемещения таблиц | Обнаружение перемещений таблиц в пределах коллекции, положения таблицы и источника изменения. | [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection), [WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs) |

## <a name="recommended-usage"></a>Рекомендуемое использование

Так `ExcelApiOnline` как API поддерживаются только Excel в Интернете, надстройка должна проверить, поддерживается ли набор требований перед вызовом этих API. Это позволяет избежать вызова API только для интернета на другой платформе.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

После того как API находится в наборе требований к платформе, необходимо удалить или изменить проверку `isSetSupported` . Это позволит включить функцию надстройки на других платформах. При внесении этого изменения обязательно проверьте эту функцию на этих платформах.

> [!IMPORTANT]
> Манифест не может указываться `ExcelApiOnline 1.1` как требование активации. Это значение не является допустимым для использования в [элементе Set](../manifest/set.md).

## <a name="api-list"></a>Список API

В следующей таблице перечислены Excel API JavaScript, включенные в набор `ExcelApiOnline` требований. Полный список всех API Excel JavaScript (`ExcelApiOnline`включая API и ранее выпущенные API), см. Excel [API JavaScript](/javascript/api/excel?view=excel-js-online&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-breaklinks-member(1))|Делает запрос на разрыв ссылок, указывающих на связанную книгу.|
||[id](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-id-member)|Исходный URL-адрес, указывающий на связанную книгу.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#excel-excel-linkedworkbook-refresh-member(1))|Делает запрос на обновление данных, извлеченных из связанной книги.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-breakalllinks-member(1))|Нарушает все ссылки на связанные книги.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitem-member(1))|Получает сведения о связанной книге по URL-адресу.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-getitemornullobject-member(1))|Получает сведения о связанной книге по URL-адресу.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-refreshall-member(1))|Делает запрос на обновление всех ссылок на книги.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#excel-excel-linkedworkbookcollection-workbooklinksrefreshmode-member)|Представляет режим обновления ссылок на книги.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-activate-member(1))|Активирует это представление листа.|
||[delete()](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-delete-member(1))|Удаляет представление листа из листа.|
||[дубликат (имя?: строка)](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-duplicate-member(1))|Создает копию этого представления листа.|
||[name](/javascript/api/excel/excel.namedsheetview#excel-excel-namedsheetview-name-member)|Получает или задает имя представления листа.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-add-member(1))|Создает новое представление листа с заданным именем.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-entertemporary-member(1))|Создает и активирует новое временное представление листа.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-exit-member(1))|Выходит из действующего представления листа.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getactive-member(1))|Получает в настоящее время активное представление листа.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getcount-member(1))|Получает количество просмотров листов в этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitem-member(1))|Получает представление листа с его именем.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-getitemat-member(1))|Получает представление листа по индексу в коллекции.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#excel-excel-namedsheetviewcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterows-member(1))|Удаление нескольких строк из таблицы.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-deleterowsat-member(1))|Удаление указанного количества строк из таблицы, начиная с указанного индекса.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedWorkbooks](/javascript/api/excel/excel.workbook#excel-excel-workbook-linkedworkbooks-member)|Возвращает коллекцию связанных книг.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-namedsheetviews-member)|Возвращает коллекцию представлений листов, присутствующих в листе.|
||[onNameChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onnamechanged-member)|Возникает при смене имени таблицы.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onvisibilitychanged-member)|Возникает при смене видимости таблицы.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onmoved-member)|Возникает при перемещении таблицы пользователем в книге.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onnamechanged-member)|Возникает при смене имени таблицы в коллекции таблиц.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onvisibilitychanged-member)|Возникает при смене видимости таблицы в коллекции таблиц.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionafter-member)|Получает новую позицию таблицы после перемещения.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionbefore-member)|Получает предыдущую позицию таблицы перед перемещением.|
||[источник](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-source-member)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-worksheetid-member)|Получает ID перемещенного таблицы.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-nameafter-member)|Получает новое имя таблицы после изменения имени.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-namebefore-member)|Получает предыдущее имя таблицы до изменения имени.|
||[источник](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-source-member)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-worksheetid-member)|Получает ID таблицы с новым именем.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[источник](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-source-member)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-type-member)|Получает тип события.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilityafter-member)|Получает новый параметр видимости таблицы после изменения видимости.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilitybefore-member)|Получает предыдущий параметр видимости таблицы перед изменением видимости.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-worksheetid-member)|Получает ID таблицы, видимость которой изменилась.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Предварительные версии API JavaScript для Excel](excel-preview-apis.md)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
