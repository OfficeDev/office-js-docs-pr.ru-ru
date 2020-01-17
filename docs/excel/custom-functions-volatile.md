---
ms.date: 01/14/2020
description: Узнайте, как реализовать переменные настраиваемые функции потоковой и автономной работы.
title: Пересчитываемые значения в функциях
localization_priority: Normal
ms.openlocfilehash: 57a41578f400b10806fc169fed09db7d7a66ce84
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217050"
---
# <a name="volatile-values-in-functions"></a>Пересчитываемые значения в функциях

Функции volatile — это функции, в которых значение изменяется каждый раз при вычислении ячейки. Значение может измениться, даже если ни один из аргументов функции не изменится. Эти функции пересчитываются при каждом пересчете в Excel. К примеру, представьте себе ячейку, вызывающую функцию `NOW`. При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`. Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).

Пользовательские функции позволяют создавать собственные переменные функции, которые могут быть удобны при обработке дат, времени, случайных чисел и моделирования. Например, для определения оптимального решения для [имитации Монте Карло](https://en.wikipedia.org/wiki/Monte_Carlo_method) требуется создание случайных входных данных.

При выборе автоматического создания JSON файла объявите переменную с помощью тега `@volatile`жсдок Comment. Дополнительные сведения об автоформировании приведены в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).

Ниже приведен пример временного настраиваемой функции, которая имитирует пошаговое описание шести костей.

![GIF-файл, в котором показана пользовательская функция, возвращающая случайное значение для имитации шести двусторонних костей](../images/six-sided-die.gif)

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

## <a name="next-steps"></a>Дальнейшие действия
Сведения о том, как [сохранить состояние в пользовательских функциях](custom-functions-save-state.md).

## <a name="see-also"></a>См. также

* [Параметры параметров пользовательских функций](custom-functions-parameter-options.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
