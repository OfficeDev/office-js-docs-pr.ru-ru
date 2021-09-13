---
title: Устранение Excel надстройки
description: Узнайте, как устранить ошибки разработки в Excel надстройки.
ms.date: 02/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 06ed12eb1daf8876e14806fd88f541b5b58eea16
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153918"
---
# <a name="troubleshooting-excel-add-ins"></a>Устранение Excel надстройки

В этой статье обсуждаются проблемы устранения неполадок, которые уникальны для Excel. Используйте средство обратной связи в нижней части страницы, чтобы предложить другие проблемы, которые можно добавить в статью.

## <a name="api-limitations-when-the-active-workbook-switches"></a>Ограничения API при активных переключателях книг

Надстройки для Excel предназначены для работы с одной книгой одновременно. Ошибки могут возникать, когда книга, которая отделена от книги, которая работает надстройка получает фокус. Это происходит только в том случае, если конкретные методы находятся в процессе призыва при смене фокуса.

На следующие API влияет этот переключатель книги.

|API JavaScript для Excel | Ошибка, брошенная |
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
> Это касается только нескольких Excel книг, открытых на Windows или Mac.

## <a name="coauthoring"></a>Совместное редактирование

См. [в](co-authoring-in-excel-add-ins.md) Excel надстройки для шаблонов, которые можно использовать с событиями в среде совместной работы. В статье также обсуждаются потенциальные конфликты слияния при использовании определенных API, например [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add_index__values_) .

## <a name="known-issues"></a>Известные проблемы

### <a name="binding-events-return-temporary-binding-obects"></a>События привязки возвращают `Binding` временные obects

Оба [bindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) и [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) возвращают временный объект, содержащий ID объекта, который поднял `Binding` `Binding` событие. Используйте этот ID для `BindingCollection.getItem(id)` получения `Binding` объекта, который поднял событие.

В следующем примере кода показано, как использовать этот временный код привязки для получения связанного `Binding` объекта. В примере слушателю событий назначена привязка. При запуске события слушатель вызывает `getBindingId` `onDataChanged` метод. Метод использует ID временного объекта для получения объекта, `getBindingId` `Binding` который поднял `Binding` событие.

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

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>Формат `useStandardHeight` ячейки `useStandardWidth` и проблемы

Свойство [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) не работает должным образом `CellPropertiesFormat` в Excel в Интернете. Из-за проблемы в пользовательском интерфейсе Excel в Интернете, задав свойство для нечетких расчетов высоты `useStandardHeight` `true` на этой платформе. Например, стандартная высота **14** изменена до **14,25** в Excel в Интернете.

На всех платформах свойства [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) и [UseStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) предназначены только для `CellPropertiesFormat` `true` . Настройка этих свойств не `false` влияет. 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Метод `getImage` Range, неподтвердимый Excel для Mac

Метод [Range getImage](/javascript/api/excel/excel.range#getImage__) в настоящее время не поддерживается в Excel для Mac. См. [в #235 OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues/235) Issue для текущего состояния.

### <a name="range-return-character-limit"></a>Ограничение возвращаемого символа диапазона

Методы [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) и [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) имеют ограничение строки адресов в 8192 символа. При превышении этого ограничения строка адресов будет усечена до 8192 символов.

## <a name="see-also"></a>Дополнительные материалы

- [Устранение ошибок разработки в надстройках Office](../testing/troubleshoot-development-errors.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
