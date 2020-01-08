---
ms.date: 12/18/2019
description: Возвращает несколько результатов из пользовательской функции в надстройке Office Excel.
title: Возвращение нескольких результатов из пользовательской функции
localization_priority: Normal
ms.openlocfilehash: 687ffcd66cff16d92fec372a778fe94bad7b38d5
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2020
ms.locfileid: "40970382"
---
# <a name="return-multiple-results-from-your-custom-function"></a>Возвращение нескольких результатов из пользовательской функции

Вы можете получить несколько результатов из пользовательской функции, которая будет возвращена соседним ячейкам. Такое поведение называется сбросом. Когда пользовательская функция возвращает массив результатов, она называется динамической формулой массива. Более подробную информацию о формулах динамических массивов в Excel можно узнать в статье [динамические массивы и функции переданных массивов](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).

На приведенном ниже изображении показано, как функция **сортировки** переключается на соседние ячейки. Пользовательская функция также может возвращать несколько результатов, как показано ниже.

![Снимок экрана функции SORT, отображающей несколько результатов в нескольких ячейках.](../images/dynamic-array-spill.png)

Чтобы создать пользовательскую функцию, которая представляет собой формулу динамической массивов, она должна возвращать двухмерный массив значений. Если результаты изменяются на соседние ячейки, у которых уже есть значения, то в формуле отображается **#SPILL!** ошибкой. 

В приведенном ниже примере показано, как вернуть динамический массив, который переключается.

```javascript
/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [['first'], ['second'], ['third']];
}
```

В приведенном ниже примере показано, как вернуть динамический массив, который наводится вправо. 

```javascript
/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [['first', 'second', 'third']];
}
```

В приведенном ниже примере показано, как вернуть динамический массив, который будет исключаться и вправо, и вниз.

```javascript
/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRectangle() {
  return [
    ['apples', 1, 'pounds'],
    ['oranges', 3, 'pounds'],
    ['pears', 5, 'crates']
  ];
}
```

## <a name="see-also"></a>См. также

- [Динамическое массивы и переопределяющее поведение массива](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Параметры для пользовательских функций Excel](custom-functions-parameter-options.md)