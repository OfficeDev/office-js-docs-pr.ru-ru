---
title: Распространенные проблемы кодирования и неожиданное поведение платформы
description: Список проблем платформы API JavaScript для Office, часто встречающихся разработчиками.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d39c379961833cdb924628becf2c2da3f7e271b9
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924796"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a>Распространенные проблемы кодирования и неожиданное поведение платформы

В этой статье описываются аспекты API JavaScript для Office, которые могут привести к непредвиденному поведению или требуют определенных шаблонов кодирования для достижения желаемого результата. Если возникла проблема, связанная с этим списком, сообщите нам об этом с помощью формы отзыва в нижней части статьи.

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a>Общие API и API Outlook не основаны на обещаниях

[Общие API](/javascript/api/office) (которые не привязаны к определенному ведущему приложению Office) и [API Outlook](/javascript/api/outlook) используют модель программирования на основе обратных вызовов. Для взаимодействия с базовым документом Office требуется асинхронный вызов чтения или записи, указывающий обратный вызов, который должен выполняться при завершении операции. Пример этого шаблона приведен в статье [Document. getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).

Эти общие API и методы API Outlook не возвращают [обещаний](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise). Таким образом, вы не можете использовать параметр [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) , чтобы приостановить выполнение до завершения асинхронной операции. Если требуется `await` поведение, можно создать оболочку вызова метода в явно созданном обещании.

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> Справочная документация содержит реализацию [файла. getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)в оболочке для обещания.

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

## <a name="excel-range-limits"></a>Пределы диапазона Excel

Если вы создаете надстройку Excel, использующую диапазоны, учитывайте следующие ограничения размера:

- В Excel в Интернете действует ограничение в 5 МБ на размер полезных данных запросов и откликов. При превышении этого ограничения возникает ошибка `RichAPI.Error`.
- Диапазон ограничен 5 000 000 ячейками для операций Set.

Если ожидается, что вводимые пользователем данные превышают эти ограничения, обязательно проверьте данные и разделите диапазоны на несколько объектов. Кроме того, вам потребуется выполнить несколько `context.sync()` вызовов, чтобы избежать появления меньших диапазонов в пакетном режиме.

Надстройка может использовать [RangeAreas](/javascript/api/excel/excel.rangeareas) для стратегических обновлений ячеек в пределах большого диапазона. Для получения дополнительных сведений просмотрите [работу с несколькими диапазонами в](../excel/excel-add-ins-multiple-ranges.md) надстройках Excel.

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
