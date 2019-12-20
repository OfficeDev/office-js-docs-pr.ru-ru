---
title: Набор обязательных элементов API JavaScript для Excel Online
description: Сведения о наборе требований Ексцелапионлине
ms.date: 12/05/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad2a3cd627552baeb449397fa917fe10e86ebbaf
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814154"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Набор обязательных элементов API JavaScript для Excel Online

Набор `ExcelApiOnline` требований является особым набором требований, включающим функции, доступные только для Excel в Интернете. API в этом наборе обязательных элементов считаются рабочими API (не подчиняются недокументированным изменениям поведения или структурным изменениям) для Excel на веб-узле. `ExcelApiOnline`считаются API-интерфейсами Preview для других платформ (Windows, Mac, iOS) и могут не поддерживаться ни одной из этих платформ.

Если API в `ExcelApiOnline` наборе обязательных элементов поддерживаются на всех платформах, они будут добавлены к следующему набору`ExcelApi 1.[NEXT]`обязательных требований (). После того как новое требование будет общедоступным, эти API будут `ExcelApiOnline`удалены из. Это можно считать похожим процессом повышения роли API, который перемещается из бета-версии в выпуск.

> [!IMPORTANT]
> `ExcelApiOnline`является надмножеством набора последних пронумерованных требований.

> [!IMPORTANT]
> `ExcelApiOnline 1.1`— Единственная версия интерфейсов API, предназначенных только для интерактивного подключения. Это связано с тем, что Excel в Интернете всегда будет иметь одну версию, доступную для пользователей с последней версией.

## <a name="recommended-usage"></a>Рекомендуемое использование

Так `ExcelApiOnline` как API поддерживаются только в Excel в Интернете, надстройка должна проверить, поддерживается ли набор требований, прежде чем вызывать эти API. Это позволяет избежать вызова Интернет-API на другой платформе.

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

После того как API находится в наборе требований к нескольким платформам, необходимо удалить или изменить `isSetSupported` проверку. При этом функция надстройки будет включена на других платформах. При внесении изменений обязательно протестируйте эту функцию на этих платформах.

> [!IMPORTANT]
> В манифесте невозможно `ExcelApiOnline 1.1` указать требования для активации. Значение не является допустимым для использования в [элементе Set](../manifest/set.md).

## <a name="api-list"></a>Список API

В данный момент в наборе обязательных элементов `ExcelApiOnline 1.1` API для Excel в Интернете доступны следующие API.

| Класс | Поля | Описание |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[mentions](/javascript/api/excel/excel.comment#mentions)|Получает объекты (например, людей), которые упоминаются в комментариях.|
||[ричконтент](/javascript/api/excel/excel.comment#richcontent)|Получает форматированный текст комментария (например, упоминания в комментариях). Эта строка не предназначена для отображения конечным пользователям. Надстройка должна использовать эту надстройку только для анализа форматированного содержимого комментариев.|
||[Упдатементионс (Контентвисментионс: Excel. Комментричконтент)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[комментментион](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Получает или задает адрес электронной почты объекта, который упоминается в примечании.|
||[id](/javascript/api/excel/excel.commentmention#id)|Получает или задает идентификатор объекта. Это соответствует одному из идентификаторов в `CommentRichContent.richContent`файле.|
||[name](/javascript/api/excel/excel.commentmention#name)|Получает или задает имя объекта, упоминаемого в примечании.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[mentions](/javascript/api/excel/excel.commentreply#mentions)|Получает объекты (например, людей), которые упоминаются в комментариях.|
||[ричконтент](/javascript/api/excel/excel.commentreply#richcontent)|Получает форматированный текст комментария (например, упоминания в комментариях). Эта строка не предназначена для отображения конечным пользователям. Надстройка должна использовать эту надстройку только для анализа форматированного содержимого комментариев.|
||[Упдатементионс (Контентвисментионс: Excel. Комментричконтент)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[комментричконтент](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Массив, содержащий все сущности (например, люди), упомянутые в комментарии.|
||[ричконтент](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[Range](/javascript/api/excel/excel.range)|[moveTo (Дестинатионранже: строка \| Range)](/javascript/api/excel/excel.range#moveto-destinationrange-)|Перемещает значения ячеек, форматирование и формулы из текущего диапазона в конечный диапазон, заменяя старые сведения в этих ячейках.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[Аджустиндент (Amount: число)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Настраивает отступ для форматирования диапазона. Значение отступа лежит в диапазоне от 0 до 250 и измеряется в символах.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-online)
- [Предварительные версии API JavaScript для Excel](./excel-preview-apis.md)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)