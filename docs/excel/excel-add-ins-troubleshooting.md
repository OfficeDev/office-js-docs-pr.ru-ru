---
title: Устранение неполадок надстроек Excel
description: Узнайте, как устранять ошибки разработки в надстройках Excel.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409412"
---
# <a name="troubleshooting-excel-add-ins"></a>Устранение неполадок надстроек Excel

В этой статье обсуждаются проблемы, связанные с устранением неполадок, которые являются уникальными для Excel. Воспользуйтесь средством обратной связи в нижней части страницы, чтобы предложить другие проблемы, которые можно добавить в статью.

## <a name="api-limitations-when-the-active-workbook-switches"></a>Ограничения API при использовании активных переключателей книги

Надстройки для Excel предназначены для работы с одной книгой за раз. Ошибки могут возникать, если книга, отделяющая от того, где работает надстройка, получает фокус. Это происходит только в том случае, если определенные методы находятся в процессе вызова при изменении фокуса.

Этот переключатель книги влияет на следующие API:

|API JavaScript для Excel | Выдается ошибка |
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
> Это относится только к нескольким книгам Excel, открываемым в Windows или Mac.

## <a name="coauthoring"></a>Совместное редактирование

Используйте совместное [Редактирование в](co-authoring-in-excel-add-ins.md) надстройках Excel для шаблонов, используемых с событиями в среде совместной работы. В этой статье также обсуждаются потенциальные конфликты объединения при использовании определенных API, например [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) .

## <a name="see-also"></a>См. также

- [Устранение ошибок разработки надстроек Office](../testing/troubleshoot-development-errors.md)
- [Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office](../testing/testing-and-troubleshooting.md)
