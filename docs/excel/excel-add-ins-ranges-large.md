---
title: Чтение или написание в больших диапазонах с Excel API JavaScript
description: Узнайте, как читать или писать в больших диапазонах с Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 4b0398c043ba7cfc4933efe600f89281c988bbb1
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150861"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a>Чтение или написание в большом диапазоне с Excel API JavaScript

В этой статье описывается обработка чтения и записи в больших диапазонах с помощью Excel API JavaScript.

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a>Запуск отдельных операций чтения или записи для больших диапазонов

Если диапазон содержит большое количество ячеек, значений, форматов номеров или формул, возможно, невозможно выполнить операции API на этом диапазоне. API всегда делает все возможное, чтобы выполнить запрошенную операцию над диапазоном (то есть получить или записать указанные данные), но попытка выполнить операцию чтения или записи для большого диапазона может привести к ошибке API из-за чрезмерного потребления ресурсов. Чтобы избежать таких ошибок, мы рекомендуем выполнять отдельные операции чтения или записи для небольших подмножеств большого диапазона, а не пытаться выполнить одну операцию чтения или записи для большого диапазона.

Дополнительные сведения об ограничениях системы см. в разделе "Excel надстройки" ограничений ресурсов и оптимизации производительности для [Office надстройки.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)

### <a name="conditional-formatting-of-ranges"></a>Условное форматирование диапазонов

В диапазонах может применяться форматирование к отдельным ячейкам на основе условий. Дополнительные сведения об этом см. в статье [Применение условного форматирования к диапазонам Excel](excel-add-ins-conditional-formatting.md).

## <a name="see-also"></a>Дополнительные материалы

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Чтение или написание в неограниченый диапазон с Excel API JavaScript](excel-add-ins-ranges-unbounded.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
