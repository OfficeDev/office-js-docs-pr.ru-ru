---
title: Обработка ошибок с Excel API JavaScript
description: Узнайте о Excel логике обработки ошибок API JavaScript для учета ошибок во время работы.
ms.date: 02/16/2022
ms.localizationpriority: medium
ms.openlocfilehash: f826968ae5137931fc334e62c9594461850cf649
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484078"
---
# <a name="error-handling-with-the-excel-javascript-api"></a>Обработка ошибок с Excel API JavaScript

При создании надстройки с использованием API JavaScript для Excel не забудьте включить логику для обработки ошибок, возникающих в среде выполнения. Это очень важно из-за асинхронного характера API.

> [!NOTE]
> Дополнительные сведения `sync()` о методе и асинхронном характере API JavaScript Excel см. в Excel объектной модели [JavaScript](excel-add-ins-core-concepts.md) в Office надстройки.

## <a name="best-practices"></a>Рекомендации

В наших [примерах](https://github.com/OfficeDev/Office-Add-in-samples) кода [и Script Lab](../overview/explore-with-script-lab.md) фрагментах, вы заметите, `Excel.run` `catch` что каждый вызов сопровождается заявлением, чтобы поймать все ошибки, которые происходят в `Excel.run`пределах . Мы рекомендуем использовать этот шаблон, когда вы будете создавать надстройки с использованием API JavaScript для Excel.

```js
$("#run").click(() => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
      // Add your Excel JavaScript API calls here.

      // Await the completion of context.sync() before continuing.
    await context.sync();
    console.log("Finished!");
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

```

## <a name="api-errors"></a>Ошибки API

Если Excel API JavaScript не удается выполнить успешно, API возвращает объект ошибки, содержащий следующие свойства.

- **code**.  Свойство `code` сообщения об ошибке содержит строку, входящую в список `OfficeExtension.ErrorCodes` или `Excel.ErrorCodes`. Например, код ошибки InvalidReference указывает, что ссылка недопустима для указанной операции. Коды ошибок не локализованы.

- **message.** Свойство `message` сообщения об ошибке содержит сводные сведения об ошибке в локализованной строке. Сообщение об ошибке не предназначено для пользователей. Код ошибки и соответствующую бизнес-логику следует использовать для определения сообщения об ошибке, которое ваша надстройка будет отображать для пользователей.

- **debugInfo.** Если в сообщении об ошибке имеется свойство `debugInfo`, в нем содержатся дополнительные сведения, которые вы можете использовать, чтобы понять причину ошибки.

> [!NOTE]
> Если вы используете метод `console.log()` для печати сообщений об ошибках в консоль, эти сообщения будет отображаться только на сервере. Конечные пользователи не будут видеть эти сообщения об ошибке в области задач надстройки или в любом Office приложении.

## <a name="error-messages"></a>Сообщения об ошибках

В таблице ниже перечислены ошибки, которые может возвращать API.

|Код ошибки | Сообщение об ошибке | Примечания |
|:----------|:--------------|:------|
|`AccessDenied` |Вы не можете выполнить запрашиваемую операцию.| |
|`ActivityLimitReached`|Достигнут предел действий.| |
|`ApiNotAvailable`|Запрашиваемый интерфейс API недоступен.| |
|`ApiNotFound`|API, который вы пытаетесь использовать, не удалось найти. Он может быть доступен в более новой версии Excel. Дополнительные [сведения см. в Excel API JavaScript](/javascript/api/requirement-sets/excel-api-requirement-sets).| |
|`BadPassword`|Предоставленный пароль является неправильным.| |
|`Conflict`|Запрос не удалось обработать из-за конфликта.| |
|`ContentLengthRequired`|Отсутствует `Content-length` заглавная головка HTTP.| |
|`EmptyChartSeries`|Попытка операции не удалась, так как серия диаграмм пуста.| |
|`FilteredRangeConflict`|Попытка операции вызывает конфликт с фильтрованным диапазоном.| |
|`FormulaLengthExceedsLimit`|Bytecode примененной формулы превышает максимальное ограничение длины. Для Office 32-битных машин ограничение длины bytecode составляет 16384 символа. На 64-битных машинах ограничение длины bytecode составляет 32768 знаков.| Эта ошибка возникает как в Excel в Интернете, так и на рабочем столе.|
|`GeneralException`|При обработке запроса возникла внутренняя ошибка.| |
|`InactiveWorkbook`|Операция не удалась, так как несколько книг открыты, а вызванная этим API книга потеряла фокус.| |
|`InsertDeleteConflict`|Операция вставки или удаления привела к конфликту.| |
|`InvalidArgument` |Аргумент недопустим, отсутствует или имеет неправильный формат.| |
|`InvalidBinding` |Эта привязка объектов недопустима из-за предыдущих обновлений.| |
|`InvalidOperation`|Выполняемая операция недопустима для этого объекта.| |
|`InvalidOperationInCellEditMode`|Операция недоступна, пока Excel в режиме Изменить ячейку. Выход Из режима редактирования с помощью клавиш **Enter** или **Tab** или путем выбора другой ячейки, а затем попробуйте еще раз.| |
|`InvalidReference`|Эта ссылка недопустима для текущей операции.| |
|`InvalidRequest`  |Не удается обработать запрос.| |
|`InvalidSelection`|Выбранный фрагмент недопустим для этой операции.| |
|`ItemAlreadyExists`|Создаваемый ресурс уже существует.| |
|`ItemNotFound` |Запрашиваемый ресурс не существует.| |
|`MemoryLimitReached`|Предел памяти достигнут. Ваше действие не удалось завершить.| |
|`MergedRangeConflict`|Не удается выполнить операцию. Таблица не может пересекаться с другой таблицей, отчетом PivotTable, результатами запроса, объединенными ячейками или XML-картой.|
|`NonBlankCellOffSheet`|Microsoft Excel не может вставить новые ячейки, так как это отодвигает непустые ячейки с конца таблицы. Эти непустые ячейки могут казаться пустыми, но имеют пустые значения, некоторое форматирование или формулу. Удалите достаточно строк или столбцов, чтобы сделать место для того, что вы хотите вставить, а затем попробуйте еще раз.| |
|`NotImplemented`|Запрашиваемая функция не реализована.| |
|`OperationCellsExceedLimit`|Попытка операции затрагивает более 33554000 ячеек.| Если эта `TableColumnCollection.add API` ошибка вызывается, подтвердим, что в таблице нет непреднамеральных данных, но за пределами таблицы. В частности, проверьте данные в наиболее правильных столбцах таблицы. Удалите непреднамеренные данные для устранения этой ошибки. Один из способов проверки того, сколько ячеек в процессе операции необходимо выполнить следующий расчет: `(number of table rows) x (16383 - (number of table columns))`. Число 16383 — это максимальное число столбцов, Excel поддерживается. <br><br>Эта ошибка возникает только в Excel в Интернете. |
|`PivotTableRangeConflict`|Попытка операции вызывает конфликт с диапазоном PivotTable.| |
|`RangeExceedsLimit`|Количество ячейки в диапазоне превысило максимально поддерживаемый номер. Дополнительные сведения см. в статье Ограничения ресурсов и [оптимизация производительности для Office надстройки](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).| |
|`RefreshWorkbookLinksBlocked`|Операция не удалась из-за того, что пользователь не предоставил разрешение на обновление внешних ссылок на книги.| |
|`RequestAborted`|Запрос прерван во время выполнения.| |
|`RequestPayloadSizeLimitExceeded`|Размер полезной нагрузки запроса превысил предел. Дополнительные сведения см. в статье Ограничения ресурсов и [оптимизация производительности для Office надстройки](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).| Эта ошибка возникает только в Excel в Интернете.|
|`ResponsePayloadSizeLimitExceeded`|Размер полезной нагрузки отклика превысил предел. Дополнительные сведения см. в статье Ограничения ресурсов и [оптимизация производительности для Office надстройки](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).|  Эта ошибка возникает только в Excel в Интернете.|
|`ServiceNotAvailable`|Служба недоступна.| |
|`Unauthenticated` |Требуемые сведения о проверке подлинности отсутствуют или недопустимы.| |
|`UnsupportedFeature`|Операция не удалась, так как исходный таблица содержит одну или несколько неподтверченных функций.| |
|`UnsupportedOperation`|Выполняемая операция не поддерживается.| |
|`UnsupportedSheet`|Этот тип листа не поддерживает эту операцию, так как это лист Макроса или Диаграммы.| |

> [!NOTE]
> В предыдущей таблице перечислены сообщения об ошибках, с которыми вы можете столкнуться при использовании Excel API JavaScript. Если вы работаете с общим API вместо приложения Excel API JavaScript, см. в Office общие коды ошибок [API](../reference/javascript-api-for-office-error-codes.md), чтобы узнать о соответствующих сообщениях об ошибках.

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Объект OfficeExtension.Error (API JavaScript для Excel)](/javascript/api/office/officeextension.error?view=excel-js-preview&preserve-view=true)
- [Коды ошибок общего API для Office](../reference/javascript-api-for-office-error-codes.md)
