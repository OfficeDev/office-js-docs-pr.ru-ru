---
ms.date: 01/14/2020
description: Узнайте, как реализовать нестабильную и офлайновую потоковую передачу пользовательских функций.
title: Пересчитываемые значения в функциях
ms.localizationpriority: medium
ms.openlocfilehash: 90f0ecea718282ce85e7e6f2b604239c18533a9a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151509"
---
# <a name="volatile-values-in-functions"></a>Пересчитываемые значения в функциях

Летучие функции — это функции, в которых значение меняется при каждом расчете ячейки. Значение может измениться, даже если ни один из аргументов функции не изменится. Эти функции пересчитываются при каждом пересчете в Excel. К примеру, представьте себе ячейку, вызывающую функцию `NOW`. При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`. Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Настраиваемые функции позволяют создавать собственные летучие функции, которые могут быть полезны при обработке дат, времени, случайных чисел и моделирования. Например, [моделирование Монте-Карло](https://en.wikipedia.org/wiki/Monte_Carlo_method) требует генерации случайных входных данных для определения оптимального решения.

Если вы решили автогенерировать файл JSON, заявите о волатильной функции с помощью тега комментариев JSDoc. `@volatile` Дополнительные сведения об автогенерации см. в [метаданных Autogenerate JSON для пользовательских функций.](custom-functions-json-autogeneration.md)

Пример волатильной настраиваемой функции, которая имитирует развертывание шестистолковой кости.

![GIF показывает настраиваемую функцию, возвращая случайное значение для имитации прокатки шести сторон кости.](../images/six-sided-die.gif)

```JS
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided() {
  return Math.floor(Math.random() * 6) + 1;
}
```

## <a name="next-steps"></a>Следующие шаги
* Узнайте о [настраиваемом параметре функций](custom-functions-parameter-options.md).

## <a name="see-also"></a>Дополнительные материалы

* [Вручную создайте метаданные JSON для пользовательских функций](custom-functions-json.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
