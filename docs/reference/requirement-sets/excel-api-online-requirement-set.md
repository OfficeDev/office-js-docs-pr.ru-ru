---
title: Набор обязательных элементов API JavaScript для Excel Online
description: Сведения о наборе требований Ексцелапионлине.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 16c96f413424d5fc85a21419fb72cf6580c1ac18
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996531"
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

| Класс | Поля | Описание |
|:---|:---|:---|
|[Range](/javascript/api/excel/excel.range)|[Жетмержедареас ()](/javascript/api/excel/excel.range#getmergedareas--)|Возвращает объект RangeAreas, представляющий Объединенные области в этом диапазоне.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Предварительные версии API JavaScript для Excel](excel-preview-apis.md)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
