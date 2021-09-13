---
title: Работа с ячейками с Excel API JavaScript.
description: Узнайте Excel API JavaScript для ячейки и узнайте, как работать с ячейками.
ms.date: 04/16/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 74603727c5944583f55e77c75589f31ffbdffb21
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151432"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Работа с ячейками с Excel API JavaScript

В API JavaScript для Excel нет объекта или класса Cell. Вместо этого все Excel являются `Range` объектами. Отдельные ячейки в пользовательском интерфейсе Excel преобразуются в объект `Range` с одной ячейкой в API JavaScript для Excel.

Объект `Range` также может содержать несколько соразмерных ячеек. Дополнительные ячейки образуют неоконченный прямоугольник (включая отдельные строки или столбцы). Чтобы узнать о работе с ячейками, которые не являются соразмерными, см. в этой ссылке Работа с дисконтными ячейками с помощью объекта [RangeAreas.](#work-with-discontiguous-cells-using-the-rangeareas-object)

Полный список свойств и методов, поддерживаемых объектом, см. в руб. `Range` [Range Object (API JavaScript для Excel).](/javascript/api/excel/excel.range)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Работа с дисконтными ячейками с помощью объекта RangeAreas

Объект [RangeAreas](/javascript/api/excel/excel.rangeareas) позволяет надстройки выполнять операции сразу на нескольких диапазонах. Эти диапазоны могут быть состоятельными, но они не должны быть. Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>Дополнительные материалы

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Получите диапазон с помощью Excel API JavaScript](excel-add-ins-ranges-get.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
