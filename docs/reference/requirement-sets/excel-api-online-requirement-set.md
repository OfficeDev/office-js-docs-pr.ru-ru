---
title: Набор обязательных элементов API JavaScript для Excel Online
description: Сведения о наборе требований Ексцелапионлине
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: aa497ff97533ff3a414905547a949fa8430c3efe
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430816"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Набор обязательных элементов API JavaScript для Excel Online

`ExcelApiOnline`Набор требований является особым набором требований, включающим функции, доступные только для Excel в Интернете. API в этом наборе обязательных элементов считаются рабочими API (не подчиняются недокументированным изменениям поведения или структурным изменениям) для Excel в веб-приложении. `ExcelApiOnline` считаются API-интерфейсами Preview для других платформ (Windows, Mac, iOS) и могут не поддерживаться ни одной из этих платформ.

Если API в `ExcelApiOnline` наборе обязательных элементов поддерживаются на всех платформах, они будут добавлены к следующему набору обязательных требований ( `ExcelApi 1.[NEXT]` ). После того как новое требование будет общедоступным, эти API будут удалены из `ExcelApiOnline` . Это можно считать похожим процессом повышения роли API, который перемещается из бета-версии в выпуск.

> [!IMPORTANT]
> `ExcelApiOnline` является надмножеством набора последних пронумерованных требований.

> [!IMPORTANT]
> `ExcelApiOnline 1.1` — Единственная версия интерфейсов API, предназначенных только для интерактивного подключения. Это связано с тем, что Excel в Интернете всегда будет иметь одну версию, доступную для пользователей с последней версией.

## <a name="recommended-usage"></a>Рекомендуемое использование

Так как `ExcelApiOnline` API поддерживаются только в Excel в Интернете, надстройка должна проверить, поддерживается ли набор требований, прежде чем вызывать эти API. Это позволяет избежать вызова Интернет-API на другой платформе.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

После того как API находится в наборе требований к нескольким платформам, необходимо удалить или изменить `isSetSupported` проверку. При этом функция надстройки будет включена на других платформах. При внесении изменений обязательно протестируйте эту функцию на этих платформах.

> [!IMPORTANT]
> В манифесте невозможно указать `ExcelApiOnline 1.1` требования для активации. Значение не является допустимым для использования в [элементе Set](../manifest/set.md).

## <a name="api-list"></a>Список API

В данный момент в наборе обязательных элементов API для Excel в Интернете доступны следующие API `ExcelApiOnline 1.1` .

| Класс | Поля | Описание |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Задает угол, по которому текст будет ориентирован на название оси диаграммы. Значение должно быть целым числом от – 90 до 90 или целым числом 180 для вертикально ориентированного текста.|
|[пивоттаблескопедколлектион](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Получает количество сводных таблиц в коллекции.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Получает первую сводную таблицу в коллекции. Сводные таблицы в коллекции сортируются сверху вниз и слева направо, так как первая сводная таблица в коллекции является верхней левой.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Получает сводную таблицу по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Получает сводную таблицу по имени. Если сводная таблица не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/excel/excel.range)|[PivotTable (Фулликонтаинед?: Boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Возвращает ограниченную коллекцию сводных таблиц, которые перекрывают диапазон.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Предварительные версии API JavaScript для Excel](./excel-preview-apis.md)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)