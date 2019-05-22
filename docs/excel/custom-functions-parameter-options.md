---
ms.date: 05/09/2019
description: Узнайте, как использовать различные параметры в пользовательских функциях, таких как диапазоны Excel, необязательные параметры, контекст вызова и многое другое.
title: Параметры для пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: 7bf195bbae696274518966e2a24bd9819e9c3f4b
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337190"
---
# <a name="custom-functions-parameter-options"></a>Параметры параметров пользовательских функций

Настраиваемые функции можно настраивать с помощью различных параметров:
- [Необязательные параметры](#custom-functions-optional-parameters)
- [Параметры Range](#range-parameters)
- [Параметр контекста вызова](#invocation-parameter)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-optional-parameters"></a>Необязательные параметры настраиваемых функций

В то время как обычные параметры являются обязательными, необязательные параметры — нет. Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках. В приведенном ниже примере функция Add может дополнительно добавить третий номер. Эта функция отображается как `=CONTOSO.ADD(first, second, [third])` в Excel.

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third !== undefined) {
    return first + second + third;
  }
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

Если вы определяете функцию, содержащую один или несколько необязательных параметров, нужно указать, что происходит, когда необязательный параметр не задан. В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`. Если `zipCode` параметр не определен, для `98052`него устанавливается значение по умолчанию. Если параметр `dayOfWeek` не определен, ему присваивается значение Wednesday (Среда).

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} zipCode Zip code. If omitted, zipCode = 98052.
 * @param {string} dayOfWeek Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

## <a name="range-parameters"></a>Параметры Range

Настраиваемая функция может принимать диапазон данных ячейки в качестве входного параметра. Функция также может возвращать диапазон данных. Excel передает диапазон данных ячейки в виде двумерного массива.

Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel. Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`. Обратите внимание, что в метаданных JSON для этой функции для `type` свойства параметра задано значение `matrix`.

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.  
 */
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
CustomFunctions.associate("SECONDHIGHEST", secondHighest);
```

## <a name="invocation-parameter"></a>Параметр вызова

Каждая пользовательская функция автоматически передает `invocation` аргумент в качестве последнего аргумента. Этот аргумент можно использовать для получения дополнительного контекста, например адреса вызывающей ячейки. Или его можно использовать для отправки в Excel данных, например обработчика функции для [отмены функции](custom-functions-web-reqs.md#stream-and-cancel-functions). Даже если вы не объявили параметры, у настраиваемой функции есть этот параметр. Этот аргумент не отображается для пользователя в Excel. Если вы хотите использовать `invocation` пользовательскую функцию, объявите ее в качестве последнего параметра.

В следующем примере кода `invocation` контекст явно указывается для ссылки.

```js
/**
 * Add two numbers.
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

Параметр позволяет получить контекст вызывающей ячейки, который может быть полезен в некоторых сценариях, в том числе [Обнаружение адреса ячейки, которая вызывает настраиваемую функцию](#addressing-cells-context-parameter).

### <a name="addressing-cells-context-parameter"></a>Параметр контекста ячейки адресации

В некоторых случаях необходимо получить адрес ячейки, которая вызвала пользовательскую функцию. Это полезно в следующих сценариях:

- Диапазоны форматирования: используйте адрес ячейки в качестве ключа для хранения информации в [оффицерунтиме. Storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data). После этого используйте событие [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) в Excel, чтобы загрузить ключ из `OfficeRuntime.storage`.
- Отображение кэшированных значений. Если функция используется в автономном режиме, отображайте сохраненные в кэше значения из `OfficeRuntime.storage` с помощью `onCalculated`.
- Сверка: используйте адрес ячейки, чтобы найти исходную ячейку, чтобы упростить сверку при выполнении обработки.

Чтобы запросить контекст ячейки адресации в функции, необходимо использовать функцию для поиска адреса ячейки, например, в приведенном ниже примере. Сведения об адресе ячейки отображаются только в том случае, `@requiresAddress` если она помечена комментариями функции.

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
CustomFunctions.associate("GETADDRESS", getAddress);
```

По умолчанию значения, возвращаемые из функции `getAddress`, соответствуют следующему формату: `SheetName!CellNumber`. Например, если функция вызвана с листа с названием Expenses (Расходы) в ячейке B2, возвращаемым значением будет `Expenses!B2`.

## <a name="next-steps"></a>Дальнейшие действия
Сведения о том, как [сохранить состояние в пользовательских функциях](custom-functions-save-state.md) или использовать [переменные значения в пользовательских функциях](custom-functions-volatile.md).

## <a name="see-also"></a>См. также

* [Получение и обработка данных с помощью пользовательских функций](custom-functions-web-reqs.md)
* [Рекомендации по пользовательским функциям](custom-functions-best-practices.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
