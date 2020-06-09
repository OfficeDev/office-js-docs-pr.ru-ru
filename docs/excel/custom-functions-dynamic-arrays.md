---
ms.date: 05/11/2020
description: Возвращает несколько результатов из пользовательской функции в надстройке Office Excel.
title: Возвращение нескольких результатов из пользовательской функции
localization_priority: Normal
ms.openlocfilehash: e25965277fbbe1c39007f79f401bf62b25760488
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609652"
---
# <a name="return-multiple-results-from-your-custom-function"></a>Возвращение нескольких результатов из пользовательской функции

Вы можете получить несколько результатов из пользовательской функции, которая будет возвращена соседним ячейкам. Такое поведение называется сбросом. Когда пользовательская функция возвращает массив результатов, она называется динамической формулой массива. Более подробную информацию о формулах динамических массивов в Excel можно узнать в статье [динамические массивы и функции переданных массивов](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).

На приведенном ниже изображении показано, как `SORT` функция переключается на соседние ячейки. Пользовательская функция также может возвращать несколько результатов, как показано ниже.

![Снимок экрана функции SORT, отображающей несколько результатов в нескольких ячейках.](../images/dynamic-array-spill.png)

Чтобы создать пользовательскую функцию, которая представляет собой формулу динамической массивов, она должна возвращать двухмерный массив значений. Если результаты изменяются на соседние ячейки, у которых уже есть значения, то формула выведет `#SPILL!` сообщение об ошибке.

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

- [Динамическое массивы и переопределяющее поведение массива](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Параметры для пользовательских функций Excel](custom-functions-parameter-options.md)