---
title: Excel Набор требований для API javaScript только для интернета
description: Сведения о наборе требований ExcelApiOnline.
ms.date: 07/23/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c8c0282970cd384ea0e7f47762c1e24c6af6536a
ms.sourcegitcommit: 3cc8f6adee0c7c68c61a42da0d97ed5ea61be0ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53661273"
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
| Именуемые представления листа | Предоставляет программный контроль представлений таблицы для каждого пользователя. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |

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
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearcolumncriteria-columnindex-)|Очищает критерии фильтрации столбцов автофайлов.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Активирует это представление листа.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Удаляет представление листа из листа.|
||[дубликат (имя?: строка)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Создает копию этого представления листа.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Получает или задает имя представления листа.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Создает новое представление листа с заданным именем.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Создает и активирует новое временное представление листа.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Выходит из действующего представления листа.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Получает в настоящее время активное представление листа.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Получает количество просмотров листов в этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Получает представление листа с его именем.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Получает представление листа по индексу в коллекции.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Возвращает коллекцию представлений листов, присутствующих в листе.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Предварительные версии API JavaScript для Excel](excel-preview-apis.md)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
