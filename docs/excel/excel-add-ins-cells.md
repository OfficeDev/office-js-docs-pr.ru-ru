---
title: Работа с ячейками с Excel API JavaScript.
description: Узнайте Excel API JavaScript для ячейки и узнайте, как работать с ячейками.
ms.date: 04/16/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f9ce806fa9478835ddf009596315108c88c4f1b4
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744640"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Работа с ячейками с Excel API JavaScript

В API JavaScript для Excel нет объекта или класса Cell. Вместо этого все Excel являются объектами`Range`. Отдельные ячейки в пользовательском интерфейсе Excel преобразуются в объект `Range` с одной ячейкой в API JavaScript для Excel.

Объект `Range` также может содержать несколько соразмерных ячеек. Дополнительные ячейки образуют неоконченный прямоугольник (включая отдельные строки или столбцы). Чтобы узнать о работе с ячейками, которые не являются соразмерными, см. в этой ссылке Работа с дисконтными ячейками [с помощью объекта RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).

Полный список свойств `Range` и методов, поддерживаемых объектом, см. в списке [Range Object (API JavaScript для Excel)](/javascript/api/excel/excel.range).

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Работа с дисконтными ячейками с помощью объекта RangeAreas

Объект [RangeAreas](/javascript/api/excel/excel.rangeareas) позволяет надстройки выполнять операции сразу на нескольких диапазонах. Эти диапазоны могут быть состоятельными, но они не должны быть. Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Получите диапазон с помощью Excel API JavaScript](excel-add-ins-ranges-get.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
