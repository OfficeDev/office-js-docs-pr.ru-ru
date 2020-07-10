---
title: Совместное редактирование в надстройках Excel
description: Сведения для совместного редактирования книги Excel, хранящейся в OneDrive, OneDrive для бизнеса или SharePoint Online.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 4414bf64f05c29328c63d0857a6e498495712ff1
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093478"
---
# <a name="coauthoring-in-excel-add-ins"></a>Совместное редактирование в надстройках Excel  

With [coauthoring](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104), multiple people can work together and edit the same Excel workbook simultaneously. All coauthors of a workbook can see another coauthor's changes as soon as that coauthor saves the workbook. To coauthor an Excel workbook, the workbook must be stored in OneDrive, OneDrive for Business, or SharePoint Online.

> [!IMPORTANT]
> В Excel для Microsoft 365 вы увидите в левом верхнем углу функцию автосохранения. Если автосохранение включено, соавторы видят внесенные вами изменения в режиме реального времени. Учтите влияние такого поведения на макет вашей надстройки Excel. Пользователи могут выключить автосохранение с помощью переключателя в левом верхнем углу окна Excel.

## <a name="coauthoring-overview"></a>Общие сведения о совместном редактировании

Когда вы меняете содержимое книги, Excel автоматически синхронизирует эти изменения для всех соавторов. Вносить изменения в содержимое книги могут не только соавторы, но и код в надстройке Excel. Например, для диапазона задается значение "Contoso", когда в надстройке Office выполняется следующий код JavaScript:

```js
range.values = [['Contoso']];
```
После того как книга со значением "Contoso" синхронизируется для всех соавторов, новое значение диапазона будет доступно всем пользователям и надстройкам, выполняемым в той же книге.

Coauthoring only synchronizes the content within the shared workbook. Values copied from the workbook to JavaScript variables in an Excel add-in are not synchronized. For example, if your add-in stores the value of a cell (such as 'Contoso') in a JavaScript variable, and then a coauthor changes the value of the cell to 'Example', after synchronization all coauthors see 'Example' in the cell. However, the value of the JavaScript variable is still set to 'Contoso'. Furthermore, when multiple coauthors use the same add-in, each coauthor has their own copy of the variable, which is not synchronized. When you use variables that use workbook content, be sure you check for updated values in the workbook before you use the variable.

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>Использование событий для управления состоянием надстройки в памяти

Excel add-ins can read workbook content (from hidden worksheets and a setting object), and then store it in data structures such as variables. After the original values are copied into any of these data structures, coauthors can update the original workbook content. This means that the copied values in the data structures are now out of sync with the workbook content. When you build your add-ins, be sure to account for this separation of workbook content and values stored in data structures.

For example, you might build a content add-in that displays custom visualizations. The state of your custom visualizations might be saved in a hidden worksheet. When coauthors use the same workbook, the following scenario can occur:

- User A opens the document and the custom visualizations are shown in the workbook. The custom visualizations read data from a hidden worksheet (for example, the color of the visualizations is set to blue).
- User B opens the same document, and starts modifying the custom visualizations. User B sets the color of the custom visualizations to orange. Orange is saved to the hidden worksheet.
- Скрытый лист пользователя А обновляется с учетом оранжевого цвета.
- Специальные элементы визуализации пользователя А по-прежнему синие.

If you want User A's custom visualizations to respond to changes made by coauthors on the hidden worksheet, use the [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event. This ensures that changes to workbook content made by coauthors is reflected in the state of your add-in.

## <a name="caveats-to-using-events-with-coauthoring"></a>Предостережения, касающиеся использования событий для совместного редактирования

As described earlier, in some scenarios, triggering events for all coauthors provides an improved user experience. However, be aware that in some scenarios this behavior can produce poor user experiences. 

Например, обычно в сценариях проверки данных пользовательский интерфейс отображается в ответ на события. Событие [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs), описанное в предыдущем разделе, выполняется когда локальный пользователь или соавтор (удаленный) изменяет содержимое книги в пределах привязки. Если обработчик события `BindingDataChanged` отображает пользовательский интерфейс, пользователи увидят пользовательский интерфейс, который не связан с изменениями, над которыми они работали в книге, что приводит к плохому взаимодействию с пользователем. Избегайте отображения пользовательского интерфейса при использовании событий в вашей надстройке.

## <a name="see-also"></a>См. также

- [О совместном редактировании в Excel (VBA)](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [Как автосохранение влияет на надстройки и макросы (VBA)](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
