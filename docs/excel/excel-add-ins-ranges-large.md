---
title: Чтение или написание в больших диапазонах с помощью API JavaScript Excel
description: Узнайте, как читать или писать в больших диапазонах с помощью API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b7a1e54d6b516889884f777bd256df8fb663c794
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652918"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a>Чтение или написание в большом диапазоне с помощью API JavaScript Excel

В этой статье описывается, как обрабатывать чтение и запись в больших диапазонах с помощью API JavaScript Excel.

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a>Запуск отдельных операций чтения или записи для больших диапазонов

Если диапазон содержит большое количество ячеек, значений, форматов номеров или формул, возможно, невозможно выполнить операции API на этом диапазоне. API всегда делает все возможное, чтобы выполнить запрошенную операцию над диапазоном (то есть получить или записать указанные данные), но попытка выполнить операцию чтения или записи для большого диапазона может привести к ошибке API из-за чрезмерного потребления ресурсов. Чтобы избежать таких ошибок, мы рекомендуем выполнять отдельные операции чтения или записи для небольших подмножеств большого диапазона, а не пытаться выполнить одну операцию чтения или записи для большого диапазона.

Дополнительные сведения об ограничениях системы см. в разделе "Надстройки Excel" ограничения ресурсов и оптимизация производительности для [надстройок Office.](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)

### <a name="conditional-formatting-of-ranges"></a>Условное форматирование диапазонов

В диапазонах может применяться форматирование к отдельным ячейкам на основе условий. Дополнительные сведения об этом см. в статье [Применение условного форматирования к диапазонам Excel](excel-add-ins-conditional-formatting.md).

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с помощью API JavaScript Excel](excel-add-ins-cells.md)
- [Чтение или написание в неограниченый диапазон с помощью API JavaScript Excel](excel-add-ins-ranges-unbounded.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
