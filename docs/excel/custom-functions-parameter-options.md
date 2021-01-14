---
ms.date: 12/21/2020
description: Узнайте, как использовать различные параметры в пользовательских функциях, такие как диапазоны Excel, необязательные параметры, контекст вызовов и другие.
title: Параметры пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: 312046551236e96e67de6f63f3e3511aba6f50ce
ms.sourcegitcommit: 48b9c3b63668b2a53ce73f92ce124ca07c5ca68c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2020
ms.locfileid: "49735531"
---
# <a name="custom-functions-parameter-options"></a>Параметры пользовательских функций

Настраиваемые функции можно настраивать с помощью множества различных параметров.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Необязательные параметры

Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках. В следующем примере функция добавления при желании может добавить третий номер. Эта функция отображается, как `=CONTOSO.ADD(first, second, [third])` в Excel.

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
> Если для необязательного параметра не задано значение, Excel назначает ему `null` значение. Это означает, что инициализированные по умолчанию параметры в TypeScript не будут работать ожидаемым образом. Не используйте синтаксис, так как он не будет инициализироваться `function add(first:number, second:number, third=0):number` `third` до 0. Вместо этого используйте синтаксис TypeScript, как показано в предыдущем примере.

При указании функции, которая содержит один или несколько необязательных параметров, укажите, что происходит, если необязательные параметры имеют null. В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`. Если параметр `zipCode` имеет значение NULL, по умолчанию задано значение `98052` . Если параметр `dayOfWeek` имеет null, ему задана среда.

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

## <a name="range-parameters"></a>Параметры range

Пользовательская функция может принимать диапазон данных ячейки в качестве входного параметра. Функция также может возвращать диапазон данных. Excel передает диапазон данных ячейки в качестве двумерного массива.

Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel. Следующая функция принимает параметр, а синтаксис JSDOC задает свойство параметра в метаданных `values` `number[][]` `dimensionality` `matrix` JSON для этой функции. 

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

Повторяюющийся параметр позволяет пользователю ввести ряд необязательных аргументов в функцию. Когда функция вызвана, значения предоставляются в массиве для параметра. Если имя параметра заканчивается числом, число каждого аргумента увеличивается постепенно, например `ADD(number1, [number2], [number3],…)` . Это соответствует соглашению, используемого для встроенных функций Excel.

Следующая функция суммирует сумму чисел, адресов ячеей, а также диапазонов, если они введены.

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

Эта функция `=CONTOSO.ADD([operands], [operands]...)` показана в книге Excel.

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>Повторяюющийся параметр с одним значением

Повторяющийся параметр с одним значением позволяет передавать несколько одно значений. Например, пользователь может ввести ADD(1,B2,3). В следующем примере показано, как объявить один параметр значения.

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

### <a name="single-range-parameter"></a>Параметр одиночного диапазона

С технической точки000 г. один параметр диапазона не является повторяются, но он включен в него, так как объявление очень похоже на повторяющие параметры. Пользователю будет отображаться как ADD(A2:B3), где из Excel передается один диапазон. В следующем примере показано, как объявить один параметр диапазона.

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

### <a name="repeating-range-parameter"></a>Параметр повторяют диапазон

Параметр повторяют диапазон позволяет передавать несколько диапазонов или чисел. Например, пользователь может ввести ADD(5,B2,C3,8,E5:E8). Повторяющиеся диапазоны обычно заданы с типом, так как `number[][][]` они являются трехмерными матрицами. Пример см. в основном примере, в списке повторяюющихся параметров (#repeating-parameters).


### <a name="declaring-repeating-parameters"></a>Объявление повторяюющихся параметров
В Typescript указать, что параметр многомерный. Например,  `ADD(values: number[])` можно указать одномерный массив, указать двумерный массив и так `ADD(values:number[][])` далее.

В JavaScript используйте одномерные массивы, двумерные массивы и так далее для `@param values {number[]}` `@param <name> {number[][]}` большего размера.

Для JSON, от руки, убедитесь, что параметр указан как в файле JSON, а также убедитесь, что параметры `"repeating": true` помечены как `"dimensionality": matrix` .

## <a name="invocation-parameter"></a>Параметр вызовов

Каждая пользовательская функция автоматически передает аргумент в качестве последнего входного параметра, даже если она не `invocation` объявлена явным образом. Этот `invocation` параметр соответствует объекту [Invocation.](/javascript/api/custom-functions-runtime/customfunctions.invocation) Объект можно использовать для получения дополнительного контекста, например адреса ячейки, которая `Invocation` вызывалась пользовательской функцией. Чтобы получить доступ `Invocation` к объекту, необходимо объявить `invocation` как последний параметр в пользовательской функции. 

> [!NOTE]
> Параметр не появляется в Excel в качестве `invocation` аргумента пользовательской функции для пользователей.

В следующем примере показано, как использовать параметр для возврата адреса ячейки, которая `invocation` вызывает пользовательскую функцию. В этом примере [используется свойство address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) `Invocation` объекта. Чтобы получить доступ `Invocation` к объекту, сначала `CustomFunctions.Invocation` объявите в качестве параметра в JSDoc. Затем `@requiresAddress` объявите в JSDoc доступ к `address` свойству `Invocation` объекта. Наконец, в функции извлекаете и возвращаете `address` свойство. 

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

В Excel пользовательская функция, вызываемая свойство объекта, возвращает абсолютный адрес, следующий за форматом в ячейке, вызываемой `address` `Invocation` `SheetName!RelativeCellAddress` функцией. Например, если входной параметр находится на листе с и названием **"Цены"** в ячейке F6, возвращается значение адреса параметра `Prices!F6` . 

Этот `invocation` параметр также можно использовать для отправки сведений в Excel. Дополнительные [дополнительные данные см.](custom-functions-web-reqs.md#make-a-streaming-function) в подкатовной функции "Сделать потоковой передачей".

## <a name="detect-the-address-of-a-parameter"></a>Обнаружение адреса параметра

В сочетании с [параметром вызовов](#invocation-parameter)можно использовать объект ["Вызов"](/javascript/api/custom-functions-runtime/customfunctions.invocation) для получения адреса параметра ввода пользовательской функции. При вызове свойство [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) объекта позволяет функции возвращать адреса `Invocation` всех входных параметров. 

Это полезно в сценариях, где типы входных данных могут отличаться. Адрес входного параметра можно использовать для проверки формата номера входного значения. Формат номера можно при необходимости скорректировать до ввода. Адрес входного параметра также можно использовать, чтобы определить, есть ли у входного значения какие-либо связанные свойства, которые могут быть релевантны для последующих вычислений. 

>[!IMPORTANT]
> В `parameterAddresses` настоящее время свойство работает только с [вручную созданными метаданными JSON.](custom-functions-json.md) Чтобы вернуть адреса параметров, у объекта должно быть задано свойство , а для объекта `options` `requiresParameterAddresses` должно быть `true` `result` `dimensionality` задано свойство `matrix` .

Следующая пользовательская функция принимает три входных параметра, извлекает свойство объекта для каждого параметра и возвращает `parameterAddresses` `Invocation` адреса. 

```js
/**
 * Return the address of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

При запуске настраиваемой функции, вызываемой свойством, адрес параметра возвращается в соответствии с форматом в ячейке, `parameterAddresses` `SheetName!RelativeCellAddress` вызываемой функцией. Например, если входной параметр расположен на листе **"Затраты"** в ячейке D8, возвращается значение адреса параметра `Costs!D8` . Если настраиваемая функция имеет несколько параметров и возвращается несколько адресов параметров, возвращаемые адреса будут перетекать по нескольким ячейкам, убывая по вертикали из ячейки, которая вызывает функцию. 

## <a name="next-steps"></a>Дальнейшие действия

Узнайте, как использовать [переменные значения в пользовательских функциях.](custom-functions-volatile.md)

## <a name="see-also"></a>См. также

* [Получение и обработка данных с помощью пользовательских функций](custom-functions-web-reqs.md)
* [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Создание метаданных JSON для пользовательских функций вручную](custom-functions-json.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
