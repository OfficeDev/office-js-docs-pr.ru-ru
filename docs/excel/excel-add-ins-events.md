---
title: Работа с событиями при помощи API JavaScript для Excel
description: Список событий для Excel JavaScript. Это включает сведения об использовании обработчиков событий и связанных шаблонов.
ms.date: 12/06/2021
ms.localizationpriority: medium
---

# <a name="work-with-events-using-the-excel-javascript-api"></a>Работа с событиями при помощи API JavaScript для Excel

В этой статье описываются важные понятия, относящиеся к работе с событиями в Excel, а также представлены образцы кода, иллюстрирующие регистрацию, использование и удаление обработчиков событий при помощи API JavaScript для Excel.

## <a name="events-in-excel"></a>События в Excel

Каждый раз, когда в книге Excel происходят изменения определенного типа, срабатывает уведомление о событии. С помощью API JavaScript для Excel можно регистрировать обработчики событий, позволяющие надстройке автоматически выполнять специальную функцию при возникновении определенного события. Ниже перечислены поддерживаемые в настоящее время события.

| Событие | Описание | Поддерживаемые объекты |
|:---------------|:-------------|:-----------|
| `onActivated` | Возникает при активации объекта. | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member), [**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member) |
| `onActivated` | Возникает при активации книги. | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member) |
| `onAdded` | Возникает при добавлении объекта в коллекцию. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member) |
| `onAutoSaveSettingChanged` | Возникает при изменении параметра `autoSave` для книги. | [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member) |
| `onCalculated` | Возникает после завершения вычислений на листе (или на всех листах коллекции). | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member) |
| `onChanged` | Происходит, когда изменились данные отдельных ячеек или комментариев. | [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member), [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member) [](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member) |
| `onColumnSorted` | Возникает при сортировке одного или нескольких столбцов. Происходит в результате операции сортировки слева направо. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member) |
| `onDataChanged` | Возникает при изменении данных или форматирования в привязке. | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-ondatachanged-member) |
| `onDeactivated` | Возникает при отключении объекта. | [**Chart**](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member), [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member), [**Shape**](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member) |
| `onDeleted` | Возникает при удалении объекта из коллекции. | [**ChartCollection**](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member), [**CommentCollection**](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member) |
| `onFormatChanged` | Возникает при изменении формата на листе. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member) |
| `onFormulaChanged` | Возникает при смене формулы. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member) |
| `onProtectionChanged` | Возникает при смене состояния защиты таблицы. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member) |
| `onRowHiddenChanged` | Возникает при изменении состояния скрытия строки на определенном листе. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowhiddenchanged-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowhiddenchanged-member) |
| `onRowSorted` | Возникает при сортировке одной или нескольких строк. Происходит в результате операции сортировки сверху вниз. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member) |
| `onSelectionChanged` | Возникает при изменении активной ячейки или выбранного диапазона. | [**Binding**](/javascript/api/excel/excel.binding#excel-excel-binding-onselectionchanged-member), [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member), [**Workbook**](/javascript/api/excel/excel.workbook#excel-excel-workbook-onselectionchanged-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)[**, WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member) |
| `onSettingsChanged` | Возникает при изменении параметров в документе. | [**SettingCollection**](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member) |
| `onSingleClicked` | Возникает, когда происходит щелчок левой кнопкой мыши или нажатие на листе. | [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member) |

### <a name="events-in-preview"></a>События в предварительной версии

> [!NOTE]
> Следующие события в настоящее время доступны только в общедоступной предварительной версии. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

| Событие | Описание | Поддерживаемые объекты |
|:---------------|:-------------|:-----------|
| `onFiltered` | Возникает при применении фильтра к объекту. | [**Table**](/javascript/api/excel/excel.table#excel-excel-table-onfiltered-member), [**TableCollection**](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onfiltered-member), [**Worksheet**](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onfiltered-member), [**WorksheetCollection**](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onfiltered-member) |

### <a name="event-triggers"></a>Триггеры событий

События в книге Excel могут вызываться:

- при взаимодействии пользователя с интерфейсом Excel, вносящим изменения в книгу;
- из кода (JavaScript) надстройки Office, вносящего изменения в книгу;
- из кода (макроса) надстройки VBA, вносящего изменения в книгу.

Любое изменение, соответствующее стандартному поведению Excel, вызывает соответствующие события в книге.

### <a name="lifecycle-of-an-event-handler"></a>Жизненный цикл обработчика событий

Обработчик событий создается при его регистрации надстройкой. Он удаляется при отмене его регистрации надстройкой или при обновлении, перезагрузке или закрытии надстройки. Обработчики событий не остаются в составе файла Excel или между сеансами с интернет-версией Excel.

> [!CAUTION]
> Когда объект, для которого зарегистрированы события, удаляется (например, таблица с зарегистрированным событием `onChanged`), обработчик событий больше не запускается, но остается в памяти, пока сеанс надстройки или Excel не обновится или не закроется.

### <a name="events-and-coauthoring"></a>События и совместное редактирование

Несколько человек могут работать вместе и [одновременно редактировать](co-authoring-in-excel-add-ins.md) одну книгу Excel. Для событий, которые может вызвать соавтор, в частности `onChanged`, соответствующий объект **Event** будет содержать свойство **source**, указывающее, кем было вызвано событие: локальным пользователем (`event.source = Local`) или удаленным соавтором (`event.source = Remote`).

## <a name="register-an-event-handler"></a>Регистрация обработчика событий

В приведенном ниже примере кода регистрируется обработчик события `onChanged` на листе под названием **Sample**. В этом коде указано, что при изменении данных на этом листе должна выполняться функция `handleChange`.

```js
Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    worksheet.onChanged.add(handleChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);
```

## <a name="handle-an-event"></a>Обработка событий

Как показано в предыдущем примере, при регистрации обработчика событий вы задаете функцию, которая должна выполняться при возникновении указанного события. Вы можете настроить эту функцию на выполнение любых действий, необходимых для вашего сценария. В приведенном ниже примере кода показана функция обработчика событий, которая просто записывает сведения о событии в консоль.

```js
function handleChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Change type of event: " + event.changeType);
                console.log("Address of event: " + event.address);
                console.log("Source of event: " + event.source);
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="remove-an-event-handler"></a>Удаление обработчика события

В приведенном ниже примере кода регистрируется обработчик событий `onSelectionChanged` на листе под названием **Sample** и определяется функция `handleSelectionChange`, которая будет выполняться при возникновении события. В нем также определяется функция `remove()`, которую можно впоследствии вызвать для удаления обработчика событий. Обратите внимание, `RequestContext` что для его удаления требуется использовать обработник событий. 

```js
var eventResult;

Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getItem("Sample");
    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
        .then(function () {
            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
        });
}).catch(errorHandlerFunction);

function handleSelectionChange(event)
{
    return Excel.run(function(context){
        return context.sync()
            .then(function() {
                console.log("Address of current selection: " + event.address);
            });
    }).catch(errorHandlerFunction);
}

function remove() {
    return Excel.run(eventResult.context, function (context) {
        eventResult.remove();

        return context.sync()
            .then(function() {
                eventResult = null;
                console.log("Event handler successfully removed.");
            });
    }).catch(errorHandlerFunction);
}
```

## <a name="enable-and-disable-events"></a>Включение и отключение событий

Производительность надстройки можно повысить с помощью отключения событий.
Например, вашему приложению, возможно, никогда не потребуется получать события, или оно может игнорировать события при выполнении пакетных изменений нескольких сущностей.

События включаются и отключаются на уровне [среды выполнения](/javascript/api/excel/excel.runtime).
Свойство `enableEvents` определяет, будут ли запускаться события и будут ли активироваться их обработчики.

В приведенном ниже примере кода показано, как включать и отключать события.

```js
Excel.run(function (context) {
    context.runtime.load("enableEvents");
    return context.sync()
        .then(function () {
            var eventBoolean = !context.runtime.enableEvents;
            context.runtime.enableEvents = eventBoolean;
            if (eventBoolean) {
                console.log("Events are currently on.");
            } else {
                console.log("Events are currently off.");
            }
        }).then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
