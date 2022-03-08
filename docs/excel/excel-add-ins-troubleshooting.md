---
title: Устранение Excel надстройки
description: Узнайте, как устранить ошибки разработки в Excel надстройки.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: c6a523354cc938ac9e9ba041ddb09f12142a3a58
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340794"
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

См. [в](co-authoring-in-excel-add-ins.md) Excel надстройки для шаблонов, которые можно использовать с событиями в среде совместной работы. В статье также обсуждаются потенциальные конфликты слияния при использовании определенных API, например [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1)).

## <a name="known-issues"></a>Известные проблемы

### <a name="binding-events-return-temporary-binding-obects"></a>События привязки возвращают временные `Binding` obects

Оба [bindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#excel-excel-bindingdatachangedeventargs-binding-member) и [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#excel-excel-bindingselectionchangedeventargs-binding-member) `Binding` возвращают временный объект, содержащий ID `Binding` объекта, который поднял событие. Используйте этот ID для `BindingCollection.getItem(id)` получения объекта `Binding` , который поднял событие.

В следующем примере кода показано, как использовать этот временный код привязки для получения связанного `Binding` объекта. В примере слушателю событий назначена привязка. При запуске `getBindingId` `onDataChanged` события слушатель вызывает метод. Метод `getBindingId` использует ID временного `Binding` объекта `Binding` для получения объекта, который поднял событие.

```js
async function run() {
    await Excel.run(async (context) => {
        // Retrieve your binding.
        let binding = context.workbook.bindings.getItemAt(0);
    
        await context.sync();
    
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);
        await context.sync();
    });
}

async function getBindingId(eventArgs) {
    await Excel.run(async (context) => {
        // Get the temporary binding object and load its ID. 
        let tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        let originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a>Формат ячейки `useStandardHeight` и `useStandardWidth` проблемы

Свойство [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) не `CellPropertiesFormat` работает должным образом в Excel в Интернете. Из-за проблемы в пользовательском интерфейсе Excel в Интернете `useStandardHeight` `true`, задав свойство для нечетких расчетов высоты на этой платформе. Например, стандартная высота **14** изменена до **14,25** в Excel в Интернете.

На всех платформах свойства [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member) и [UseStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member) `CellPropertiesFormat` `true`предназначены только для . Настройка этих свойств не влияет `false` .

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a>Метод Range`getImage`, неподтвердимый Excel для Mac

Метод [Range getImage](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1)) в настоящее время не поддерживается в Excel для Mac. Для [текущего состояния см. в выпуске OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) .

### <a name="range-return-character-limit"></a>Ограничение возвращаемого символа диапазона

Методы [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) и [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1)) имеют ограничение строки адресов в 8192 символа. При превышении этого ограничения строка адресов будет усечена до 8192 символов.

## <a name="see-also"></a>Дополнительные материалы

- [Устранение ошибок разработки в надстройках Office](../testing/troubleshoot-development-errors.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
