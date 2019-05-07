---
ms.date: 05/03/2019
description: Узнайте, как реализовать переменные настраиваемые функции потоковой и автономной работы.
title: Переменные значения в функциях
localization_priority: Normal
ms.openlocfilehash: 1ca3edc3de2d9ac5f2171004f89466352c5cfa1e
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2019
ms.locfileid: "33627999"
---
# <a name="volatile-values-in-functions"></a>Переменные значения в функциях

Функции volatile — это функции, в которых значение изменяется каждый раз при вычислении ячейки. Значение может измениться, даже если ни один из аргументов функции не изменится. Эти функции пересчитываются при каждом пересчете в Excel. К примеру, представьте себе ячейку, вызывающую функцию `NOW`. При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`. Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Пользовательские функции позволяют создавать собственные переменные функции, которые могут быть удобны при обработке дат, времени, случайных чисел и моделирования. Например, для определения оптимального решения для [имитации Монте Карло](https://en.wikipedia.org/wiki/Monte_Carlo_method
) требуется создание случайных входных данных.

При выборе автоматического создания JSON файла объявите переменную с помощью тега `@volatile`жсдок Comment. Дополнительные сведения об автоформировании приведены в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).

## <a name="next-steps"></a>Дальнейшие действия
Сведения о том, как [сохранить состояние в пользовательских функциях](custom-functions-save-state.md).

## <a name="see-also"></a>См. также

* [Параметры параметров пользовательских функций](custom-functions-parameter-options.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
