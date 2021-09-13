---
ms.date: 05/11/2020
description: Возвращайте несколько результатов из настраиваемой функции в Office Excel надстройки.
title: Возвращение нескольких результатов из настраиваемой функции
ms.localizationpriority: medium
ms.openlocfilehash: 63a9e74d311f26a618312a3ab7f81238cad69c21
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150893"
---
# <a name="return-multiple-results-from-your-custom-function"></a>Возвращение нескольких результатов из настраиваемой функции

Вы можете вернуть несколько результатов из настраиваемой функции, которые будут возвращены соседним ячейкам. Такое поведение называется разливом. Когда настраиваемая функция возвращает массив результатов, она называется динамической формулой массива. Дополнительные сведения о формулах динамических массивов в Excel см. в динамических массивах [и выплеснуло поведение массивов.](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)

На следующем изображении показано, `SORT` как функция передается в соседние ячейки. Ваша настраиваемая функция также может возвращать несколько результатов, как это.

![Снимок экрана функции SORT, отображающий несколько результатов вниз в несколько ячеек.](../images/dynamic-array-spill.png)

Чтобы создать настраиваемую функцию, которая является динамической формулой массива, необходимо вернуть двумерный массив значений. Если результаты перетекают в соседние ячейки, которые уже имеют значения, в формуле будет отображаться `#SPILL!` ошибка.

В следующем примере показано, как вернуть динамический массив, который выливается вниз.

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

В следующем примере показано, как вернуть динамический массив, который правильно разливается. 

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

В следующем примере показано, как вернуть динамический массив, который разливается как вниз, так и вправо.

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

## <a name="see-also"></a>Дополнительные материалы

- [Динамические массивы и поведение разлитого массива](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Параметры Excel пользовательских функций](custom-functions-parameter-options.md)