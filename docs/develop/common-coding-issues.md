---
title: Распространенные проблемы кодирования и неожиданное поведение платформы
description: Список проблем платформы API JavaScript для Office, часто встречающихся разработчиками.
ms.date: 01/02/2020
localization_priority: Normal
ms.openlocfilehash: fa33451550ab02f76a8b41ebf682e6a73d2a3a96
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969495"
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

## <a name="some-properties-cannot-be-set-directly"></a>Некоторые свойства невозможно задать напрямую

> [!NOTE]
> Этот раздел относится только к API, предназначенным для ведущего приложения, для Excel и Word.

Некоторые свойства не могут быть заданы, несмотря на то, что они доступны для записи. Эти свойства являются частью родительского свойства, которое должно быть задано как один объект. Это связано с тем, что родительское свойство использует вложенные свойства с определенными логическими связями. Эти родительские свойства должны быть заданы с помощью нотации литерала объекта, чтобы задать весь объект, а не задавать отдельные вложенные свойства этого объекта. Один из примеров этого примера находится в файле [PageLayout](/javascript/api/excel/excel.pagelayout). Свойство должно быть задано с помощью одного объекта Пажелайаутзумоптионс, как показано ниже: [](/javascript/api/excel/excel.pagelayoutzoomoptions) `zoom`

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

В предыдущем примере вы ***не*** сможете напрямую присвоить `zoom` значение: `sheet.pageLayout.zoom.scale = 200;`. Этот оператор выдает ошибку, `zoom` так как не загружен. Даже если `zoom` были загружены, набор масштабов не вступит в силу. Все операции контекста выполняются `zoom`, обновляя прокси-объект в надстройке и перезаписывая локально заданные значения.

Это поведение отличается от [свойств навигации](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) , таких как [Range. Format](/javascript/api/excel/excel.range#format). Свойства `format` можно задать с помощью навигации по объектам, как показано ниже:

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

Можно определить свойство, для которого не могут быть заданы вложенные свойства, путем проверки модификатора только для чтения. Все свойства, доступные только для чтения, могут иметь нередактируемые вложенные свойства, не предназначенные только для чтения. Записываемые свойства, `PageLayout.zoom` такие как, должны быть заданы на уровне объекта. В сводке:

- Свойство только для чтения: вложенные свойства можно задать с помощью навигации.
- Записываемое свойство: подсвойства невозможно задать с помощью навигации (необходимо задать в качестве части исходного назначения родительского объекта).

## <a name="excel-data-transfer-limits"></a>Пределы переноса данных Excel

При создании надстройки Excel учитывайте следующие ограничения размера при взаимодействии с книгой:

- В Excel в Интернете действует ограничение в 5 МБ на размер полезных данных запросов и откликов. При превышении этого ограничения возникает ошибка `RichAPI.Error`.
- Диапазон ограничен 5 000 000 ячейками для операций Get.

Если ожидается, что вводимые пользователем данные превышают эти ограничения, обязательно проверьте данные перед вызовом `context.sync()`. При необходимости разделите операцию на небольшие части. Не забудьте позвонить `context.sync()` по каждой подоперации, чтобы избежать повторного пакетной операции.

Эти ограничения обычно превышаются с помощью больших диапазонов. Надстройка может использовать [RangeAreas](/javascript/api/excel/excel.rangeareas) для стратегических обновлений ячеек в пределах большого диапазона. Для получения дополнительных сведений просмотрите [работу с несколькими диапазонами в](../excel/excel-add-ins-multiple-ranges.md) надстройках Excel.

## <a name="setting-read-only-properties"></a>Установка свойств, предназначенных только для чтения

[Определения TypeScript](referencing-the-javascript-api-for-office-library-from-its-cdn.md) для Office JS указывают, какие свойства объекта доступны только для чтения. Если вы попытаетесь установить свойство, доступное только для чтения, операция записи завершится с ошибкой без уведомления и не выдается сообщение об ошибке. В следующем примере ошибочно попытаются задать свойство, доступное только для чтения, [Chart.ID](/javascript/api/excel/excel.chart#id).

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a>Удаление обработчиков событий

Обработчики событий должны быть удалены с использованием `RequestContext` того же, в котором они были добавлены. Если надстройка должна удалить обработчик события во время выполнения, необходимо сохранить объект контекста, используемый для добавления обработчика.

```js
Excel.run(async (context) => {
    [...]

    // To later remove an event handler, store the context somewhere accessible to the handler removal function.
    // You may find it helpful to also store the event handler object and associate it with the context.
    selectionChangedHandler = myWorksheet.onSelectionChanged.add(callback);
    savedContext = currentContext;
    return context.sync();
}
```

## <a name="see-also"></a>См. также

- [OfficeDev/Office-JS](https://github.com/OfficeDev/office-js/issues): место для создания отчетов и просмотра проблем с платформой надстроек Office и API JavaScript.
- [Переполнение стека](https://stackoverflow.com/questions/tagged/office-js): место для Ask и просмотра вопросов по программированию, посвященных API JavaScript для Office. При публикации в стеке обязательно примените к вопросу тег "Office — JS".
- [UserVoice](https://officespdev.uservoice.com/): в этом месте вы можете предложить новые функции для платформы надстроек Office и API JavaScript для Office.
