---
title: Работа с ячейками с Excel API JavaScript.
description: Узнайте Excel API JavaScript для ячейки и узнайте, как работать с ячейками.
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 444feecd4aafb0e884de05b2ff198a3ca1423a16644c537865bcfb6905684a40
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079350"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Работа с ячейками с Excel API JavaScript

В API JavaScript для Excel нет объекта или класса Cell. Вместо этого все Excel являются `Range` объектами. Отдельные ячейки в пользовательском интерфейсе Excel преобразуются в объект `Range` с одной ячейкой в API JavaScript для Excel.

Объект `Range` также может содержать несколько соразмерных ячеек. Дополнительные ячейки образуют неоконченный прямоугольник (включая отдельные строки или столбцы). Чтобы узнать о работе с ячейками, которые не являются соразмерными, см. в этой ссылке Работа с дисконтными ячейками с помощью объекта [RangeAreas.](#work-with-discontiguous-cells-using-the-rangeareas-object)

Полный список свойств и методов, поддерживаемых объектом, см. в руб. `Range` [Range Object (API JavaScript для Excel).](/javascript/api/excel/excel.range)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Работа с дисконтными ячейками с помощью объекта RangeAreas

Объект [RangeAreas](/javascript/api/excel/excel.rangeareas) позволяет надстройки выполнять операции сразу на нескольких диапазонах. Эти диапазоны могут быть состоятельными, но они не должны быть. Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Получите диапазон с помощью Excel API JavaScript](excel-add-ins-ranges-get.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
