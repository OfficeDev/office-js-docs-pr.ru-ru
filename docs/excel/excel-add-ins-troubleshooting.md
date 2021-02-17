---
title: Устранение неполадок надстройки Excel
description: Узнайте, как устранять ошибки разработки в надстройки Excel.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 0efc8b4d25d9d748975146e187104972e4ad58a9
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270730"
---
# <a name="troubleshooting-excel-add-ins"></a>Устранение неполадок надстройки Excel

В этой статье обсуждается устранение неполадок, уникальных для Excel. Используйте средство обратной связи в нижней части страницы, чтобы предложить другие проблемы, которые можно добавить в статью.

## <a name="api-limitations-when-the-active-workbook-switches"></a>Ограничения API при переключении активной книги

Надстройки для Excel предназначены для одновременной работы с одной книгой. Ошибки могут возникать, когда книга, отделенная от книги, на которую запущена надстройка, получает фокус. Это происходит только в том случае, если конкретные методы находятся в процессе, когда фокус изменяется.

Этот переключатель книги влияет на следующие API::

|API JavaScript для Excel | Ошибка |
|--|--|
| `Chart.activate` | GeneralException |
| `Range.select` | GeneralException |
| `Table.clearFilters` | GeneralException |
| `Workbook.getActiveCell`  | InvalidSelection|
| `Workbook.getSelectedRange` | InvalidSelection|
| `Workbook.getSelectedRanges`  | InvalidSelection|
| `Worksheet.activate` | GeneralException |
| `Worksheet.delete`  | InvalidSelection|
| `Worksheet.gridlines` | GeneralException |
| `Worksheet.showHeadings` | GeneralException |
| `WorksheetCollection.add` | GeneralException |
| `WorksheetFreezePanes.freezeAt` | GeneralException |
| `WorksheetFreezePanes.freezeColumns` | GeneralException |
| `WorksheetFreezePanes.freezeRows` | GeneralException |
| `WorksheetFreezePanes.getLocationOrNullObject`| GeneralException |
| `WorksheetFreezePanes.unfreeze` | GeneralException |

> [!NOTE]
> Это относится только к нескольким книгам Excel, открытым в Windows или Mac.

## <a name="coauthoring"></a>Совместное редактирование

Шаблоны для использования с событиями в среде совместной работы см. в надстройках [Excel.](co-authoring-in-excel-add-ins.md) В этой статье также обсуждаются потенциальные конфликты слияния при использовании определенных API, например [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .

## <a name="known-issues"></a>Известные проблемы

### <a name="binding-events-return-temporary-binding-obects"></a>События привязки возвращают `Binding` временные обтекания

[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) и [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) возвращают временный объект, содержащий ИД объекта, который вызывает `Binding` `Binding` событие. Используйте этот ИД для `BindingCollection.getItem(id)` получения `Binding` объекта, который вызывает событие.

В следующем примере кода показано, как использовать этот временный ИД привязки для получения связанного `Binding` объекта. В примере прослушиватель событий назначен привязке. Прослушиватель вызывает метод `getBindingId` при `onDataChanged` запуске события. Метод использует ИД временного объекта для извлечения объекта, который `getBindingId` `Binding` вызывает `Binding` событие.

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>Формат и `useStandardHeight` `useStandardWidth` проблемы в ячейках

Свойство [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) не работает должным образом `CellPropertiesFormat` в Excel в Интернете. Из-за проблемы в пользовательском интерфейсе Excel в Интернете установка свойства для некорректного вычисления высоты `useStandardHeight` `true` на этой платформе. Например, стандартная высота **14** в Excel в Интернете изменена на **14,25.**

На всех платформах свойства [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) и [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) предназначены только для `CellPropertiesFormat` `true` этого. Установка этих свойств не `false` оказывает влияния. 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Метод Range `getImage` неподтверчен в Excel для Mac

Метод Range [getImage](/javascript/api/excel/excel.range#getImage__) в настоящее время не поддерживается в Excel для Mac. Текущее состояние см. в #235 [officeDev/office-js Issue.](https://github.com/OfficeDev/office-js/issues/235)

### <a name="range-return-character-limit"></a>Ограничение возвращаемого диапазона символов

Для [методов Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) и [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) ограничение строк адресов составляет 8192 символа. При превышении этого ограничения строка адреса усечена до 8192 символов.

## <a name="see-also"></a>См. также

- [Устранение ошибок разработки с помощью надстройки Office](../testing/troubleshoot-development-errors.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
