---
ms.date: 07/15/2019
description: Узнайте, как использовать различные параметры в пользовательских функциях, таких как диапазоны Excel, необязательные параметры, контекст вызова и многое другое.
title: Параметры для пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: 66e873117b82ed7258b5965a6e964f4b9e01df21
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719485"
---
# <a name="custom-functions-parameter-options"></a>Параметры параметров пользовательских функций

Настраиваемые функции можно настраивать с помощью различных параметров.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Необязательные параметры

В то время как обычные параметры являются обязательными, необязательные параметры — нет. Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках. В приведенном ниже примере функция Add может дополнительно добавить третий номер. Эта функция отображается как `=CONTOSO.ADD(first, second, [third])` в Excel.

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

---

> [!NOTE]
> Если для необязательного параметра не указано значение, Excel присваивает ему значение `null`. Это означает, что параметры, инициализированные по умолчанию в TypeScript, не будут работать должным образом. Поэтому не следует использовать синтаксис `function add(first:number, second:number, third=0):number` , так как он не инициализируется `third` до 0. Вместо этого используйте синтаксис TypeScript, как показано в предыдущем примере.

При определении функции, которая содержит один или несколько необязательных параметров, следует указать, что происходит, если необязательные параметры имеют значение null. В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`. Если `zipCode` параметр имеет значение null, для `98052`него устанавливается значение по умолчанию. Если `dayOfWeek` параметр имеет значение null, ему присваивается значение среда.

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## <a name="range-parameters"></a>Параметры Range

Настраиваемая функция может принимать диапазон данных ячейки в качестве входного параметра. Функция также может возвращать диапазон данных. Excel передает диапазон данных ячейки в виде двумерного массива.

Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel. Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`. Обратите внимание, что в метаданных JSON для этой функции для `type` свойства параметра задано значение `matrix`.

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="repeating-parameters"></a>Повторяющиеся параметры

Повторяющийся параметр позволяет пользователю ввести ряд необязательных аргументов функции. При вызове функции значения задаются в массиве для параметра. Если имя параметра заканчивается числом, каждый аргумент увеличит значение, например `ADD(number1, [number2], [number3],…)`. Это соответствует соглашению, используемому для встроенных функций Excel.

Приведенная ниже функция суммирует сумму чисел, адресов ячеек, а также диапазонов, если они введены.

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function ADD(operands: number[][][]): number {
  let total: number = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
```

Эта функция отображается `=CONTOSO.ADD([operands], [operands]...)` в книге Excel.

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>Повторяющийся параметр с одним значением

Повторяющийся одиночный параметр значения позволяет передавать несколько отдельных значений. Например, пользователь может ввести ADD (1, B2, 3). В следующем примере показано, как объявить параметр с одним значением.

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### <a name="single-range-parameter"></a>Один параметр Range

Один параметр диапазона технически не является повторяющимся параметром, но включается здесь, так как объявление очень похоже на повторяющиеся параметры. Она будет выглядеть как ADD (a2: B3), где один диапазон передается из Excel. В следующем примере показано, как объявить один параметр Range.

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### <a name="repeating-range-parameter"></a>Параметр повторяющегося диапазона

Параметр повторяющегося диапазона позволяет передавать несколько диапазонов или номеров. Например, пользователь может ввести ADD (5, B2, C3, 8, No5: E8). Повторяющиеся диапазоны обычно указываются с `number[][][]` типом, так как они представляют собой трехмерные матрицы. Пример приведен в основном примере для повторяющихся параметров (#repeating-Parameters).


### <a name="declaring-repeating-parameters"></a>Объявление повторяющихся параметров
В typescript укажите, что параметр является многомерным. Например, `ADD(values: number[])` указывает на одномерный массив, `ADD(values:number[][])` который указывает на двухмерный массив и т. д.

В JavaScript используйте `@param values {number[]}` одномерные массивы, `@param <name> {number[][]}` для двумерных массивов и т. д. для дополнительных измерений.

Для созданного вручную JSON убедитесь, что параметр указан как `"repeating": true` в файле JSON, а также проверьте, что параметры помечены как. `"dimensionality": matrix`

>[!NOTE]
>Функции, содержащие повторяющиеся параметры, автоматически содержат параметр вызова в качестве последнего параметра. Дополнительные сведения о параметрах вызова можно найти в следующем разделе.

## <a name="invocation-parameter"></a>Параметр вызова

Каждая пользовательская функция автоматически передает `invocation` аргумент в качестве последнего аргумента. Этот аргумент можно использовать для получения дополнительного контекста, например адреса вызывающей ячейки. Или его можно использовать для отправки в Excel данных, например обработчика функции для [отмены функции](custom-functions-web-reqs.md#make-a-streaming-function). Даже если вы не объявили параметры, у настраиваемой функции есть этот параметр. Этот аргумент не отображается для пользователя в Excel. Если вы хотите использовать `invocation` пользовательскую функцию, объявите ее в качестве последнего параметра.

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
```

Параметр позволяет получить контекст вызывающей ячейки, который может быть полезен в некоторых сценариях, в том числе [Обнаружение адреса ячейки, которая вызывает настраиваемую функцию](#addressing-cells-context-parameter).

### <a name="addressing-cells-context-parameter"></a>Параметр контекста ячейки адресации

В некоторых случаях необходимо получить адрес ячейки, которая вызвала пользовательскую функцию. Это полезно в следующих сценариях:

- Диапазоны форматирования: используйте адрес ячейки в качестве ключа для хранения информации в [оффицерунтиме. Storage](../excel/custom-functions-runtime.md#storing-and-accessing-data). После этого используйте событие [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) в Excel, чтобы загрузить ключ из `OfficeRuntime.storage`.
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
```

По умолчанию значения, возвращаемые из функции `getAddress`, соответствуют следующему формату: `SheetName!CellNumber`. Например, если функция вызвана с листа с названием Expenses (Расходы) в ячейке B2, возвращаемым значением будет `Expenses!B2`.

## <a name="next-steps"></a>Дальнейшие действия

Сведения о том, как [сохранить состояние в пользовательских функциях](custom-functions-save-state.md) или использовать [переменные значения в пользовательских функциях](custom-functions-volatile.md).

## <a name="see-also"></a>См. также

* [Получение и обработка данных с помощью пользовательских функций](custom-functions-web-reqs.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)