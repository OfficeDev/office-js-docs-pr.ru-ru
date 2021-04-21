---
title: Работа с ячейками с помощью API JavaScript Excel.
description: Узнайте определение API JavaScript Excel для ячейки и узнайте, как работать с ячейками.
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8ca985b6bbdcf19920c36c371e690f61639f16
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917102"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Работа с ячейками с помощью API JavaScript Excel

В API JavaScript для Excel нет объекта или класса Cell. Вместо этого все ячейки Excel являются `Range` объектами. Отдельные ячейки в пользовательском интерфейсе Excel преобразуются в объект `Range` с одной ячейкой в API JavaScript для Excel.

Объект `Range` также может содержать несколько соразмерных ячеек. Дополнительные ячейки образуют неоконченный прямоугольник (включая отдельные строки или столбцы). Чтобы узнать о работе с ячейками, которые не являются соразмерными, см. в этой ссылке Работа с дисконтными ячейками с помощью объекта [RangeAreas.](#work-with-discontiguous-cells-using-the-rangeareas-object)

Полный список свойств и методов, поддерживаемых объектом, см. в списке `Range` [Range Object (API JavaScript для Excel).](/javascript/api/excel/excel.range)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Работа с дисконтными ячейками с помощью объекта RangeAreas

Объект [RangeAreas](/javascript/api/excel/excel.rangeareas) позволяет надстройки выполнять операции сразу на нескольких диапазонах. Эти диапазоны могут быть состоятельными, но они не должны быть. Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Получите диапазон с помощью API JavaScript Excel](excel-add-ins-ranges-get.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
