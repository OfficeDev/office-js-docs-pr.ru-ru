---
title: Работа с ячейками с помощью API JavaScript Excel.
description: Узнайте определение API JavaScript Excel для ячейки и узнайте, как работать с ячейками.
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652978"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Работа с ячейками с помощью API JavaScript Excel

API JavaScript Excel не имеет объекта или класса Cell. Вместо этого все ячейки Excel являются `Range` объектами. Индивидуальная ячейка в пользовательском интерфейсе Excel преобразуется в объект с одной ячейкой в `Range` API JavaScript Excel.

Объект `Range` также может содержать несколько соразмерных ячеек. Дополнительные ячейки образуют неоконченный прямоугольник (включая отдельные строки или столбцы). Чтобы узнать о работе с ячейками, которые не являются соразмерными, см. в этой ссылке Работа с дисконтными ячейками с помощью объекта [RangeAreas.](#work-with-discontiguous-cells-using-the-rangeareas-object)

Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)

## <a name="excel-javascript-apis-that-mention-cells"></a>API Excel JavaScript, в которых упоминаются ячейки

Несмотря на то, что API JavaScript Excel не имеет объекта или класса "Cell", в ряде имен API упоминаются ячейки. Эти API контролируют свойства ячейки, такие как цвет, форматирование текста и шрифт.

Следующий список API JavaScript Excel относится к ячейкам.

- [CellBorder](/javascript/api/excel/excel.cellborder)
- [CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)
- [CellProperties](/javascript/api/excel/excel.cellproperties)
- [CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)
- [CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)
- [CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)
- [CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)
- [CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)
- [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)
- [SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Работа с дисконтными ячейками с помощью объекта RangeAreas

Объект [RangeAreas](/javascript/api/excel/excel.rangeareas) позволяет надстройки выполнять операции сразу на нескольких диапазонах. Эти диапазоны могут быть состоятельными, но они не должны быть. Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Получите диапазон с помощью API JavaScript Excel](excel-add-ins-ranges-get.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
