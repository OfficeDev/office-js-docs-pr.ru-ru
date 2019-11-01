---
title: Распространенные проблемы кодирования и неожиданное поведение платформы
description: Список проблем платформы API JavaScript для Office, часто встречающихся разработчиками.
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 8cea95e3214585ba8e0b77535916f9c564dde9df
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902200"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a>Распространенные проблемы кодирования и неожиданное поведение платформы

В этой статье описываются аспекты API JavaScript для Office, которые могут привести к непредвиденному поведению или требуют определенных шаблонов кодирования для достижения желаемого результата. Если возникла проблема, связанная с этим списком, сообщите нам об этом с помощью формы отзыва в нижней части статьи.

## <a name="some-properties-must-be-set-with-json-structs"></a>Некоторые свойства должны быть заданы с помощью структуры JSON

> [!NOTE]
> Этот раздел относится только к API, предназначенным для ведущего приложения, для Excel и Word.

Некоторые свойства должны быть заданы как структуры JSON, а не как задавать отдельные вложенные свойства. Один из примеров этого примера находится в файле [PageLayout](/javascript/api/excel/excel.pagelayout). Свойство должно быть задано с помощью одного объекта Пажелайаутзумоптионс, как показано ниже: [](/javascript/api/excel/excel.pagelayoutzoomoptions) `zoom`

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

В предыдущем примере вы ***не*** сможете напрямую присвоить `zoom` значение: `sheet.pageLayout.zoom.scale = 200;`. Этот оператор выдает ошибку, `zoom` так как не загружен. Даже если `zoom` были загружены, набор масштабов не вступит в силу. Все операции контекста выполняются `zoom`, обновляя прокси-объект в надстройке и перезаписывая локально заданные значения.

Это поведение отличается от [свойств навигации](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , таких как [Range. Format](/javascript/api/excel/excel.range#format). Свойства `format` можно задать с помощью навигации по объектам, как показано ниже:

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Можно определить свойство, для которого должны быть заданы вложенные свойства структуры JSON, путем проверки модификатора "только чтение". Все свойства, доступные только для чтения, могут иметь нередактируемые вложенные свойства, не предназначенные только для чтения. Записываемые свойства, `PageLayout.zoom` такие как, должны быть заданы с помощью структуры JSON. В сводке:

- Свойство только для чтения: вложенные свойства можно задать с помощью навигации.
- Записываемое свойство: вложенные свойства должны быть заданы с помощью структуры JSON (и не могут быть заданы с помощью навигации).

## <a name="setting-read-only-properties"></a>Установка свойств, предназначенных только для чтения

[Определения TypeScript](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) для Office JS указывают, какие свойства объекта доступны только для чтения. Если вы попытаетесь установить свойство, доступное только для чтения, операция записи завершится с ошибкой без уведомления и не выдается сообщение об ошибке. В следующем примере ошибочно попытаются задать свойство, доступное только для чтения, [Chart.ID](/javascript/api/excel/excel.chart#id).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a>См. также

- [OfficeDev/Office-JS](https://github.com/OfficeDev/office-js/issues): место для создания отчетов и просмотра проблем с платформой надстроек Office и API JavaScript.
- [Переполнение стека](https://stackoverflow.com/questions/tagged/office-js): место для Ask и просмотра вопросов по программированию, посвященных API JavaScript для Office. При публикации в стеке обязательно примените к вопросу тег "Office — JS".
- [UserVoice](https://officespdev.uservoice.com/): в этом месте вы можете предложить новые функции для платформы надстроек Office и API JavaScript для Office.
