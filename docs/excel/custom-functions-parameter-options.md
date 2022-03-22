---
ms.date: 03/08/2021
description: Узнайте, как использовать различные параметры в настраиваемой функции, такие как диапазоны Excel, необязательные параметры, контекст вызовов и другие.
title: Параметры Excel пользовательских функций
ms.localizationpriority: medium
ms.openlocfilehash: 2cc0c825932afe3a70d0f9ab6483327051c199fd
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711023"
---
# <a name="custom-functions-parameter-options"></a>Параметры настраиваемой функции

Настраиваемые функции настраиваются с различными параметрами параметров.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Необязательные параметры

Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках. В следующем примере функция добавления может дополнительно добавить третий номер. Эта функция отображается как в `=CONTOSO.ADD(first, second, [third])` Excel.

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
> Если для необязательных параметров не указывается значение, Excel назначает ему значение `null`. Это означает, что инициализированные по умолчанию параметры в TypeScript будут работать не так, как ожидалось. Не используйте синтаксис `function add(first:number, second:number, third=0):number` , так как он не будет инициализировать до `third` 0. Вместо этого используйте синтаксис TypeScript, как показано в предыдущем примере.

Когда вы определяете функцию, которая содержит один или несколько необязательных параметров, укажите, что происходит, когда необязательные параметры являются null. В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`. Если параметр `zipCode` null, значение по умолчанию задано `98052`. Если параметр `dayOfWeek` null, он задан в среду.

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

## <a name="range-parameters"></a>Параметры диапазона

Ваша настраиваемая функция может принимать ряд данных ячейки в качестве параметра ввода. Функция также может возвращать ряд данных. Excel будет передавать ряд данных ячейки в качестве двухмерного массива.

Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel. Следующая функция принимает `values`параметр, и синтаксис JSDOC `number[][]` `dimensionality` `matrix` задает свойство параметра в метаданных JSON для этой функции. 

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

## <a name="repeating-parameters"></a>Повторяющие параметры

Параметр повторения позволяет пользователю ввести ряд необязательных аргументов в функцию. Когда функция называется, значения предоставляются в массиве для параметра. Если имя параметра заканчивается числом, число каждого аргумента будет увеличиваться постепенно, например `ADD(number1, [number2], [number3],…)`. Это соответствует конвенции, используемой для встроенных Excel функций.

В следующей функции суммируется общее число, адреса ячейки, а также диапазоны, если они введены.

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

Эта функция показана `=CONTOSO.ADD([operands], [operands]...)` в Excel книге.

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>Повторение параметра единого значения

Повторяющийся параметр единого значения позволяет передать несколько одиночных значений. Например, пользователь может ввести ADD(1,B2,3). В следующем примере показано, как объявить один параметр значения.

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

### <a name="single-range-parameter"></a>Параметр "Один диапазон"

Параметр одного диапазона технически не является параметром повторения, но включается здесь, так как объявление очень похоже на повторяющие параметры. Он будет отображаться пользователю как ADD (A2:B3), где один диапазон передается из Excel. В следующем примере показано, как объявить один параметр диапазона.

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

### <a name="repeating-range-parameter"></a>Параметр "Повторение диапазона"

Параметр диапазона повторяемого диапазона позволяет передавать несколько диапазонов или номеров. Например, пользователь может ввести ADD(5,B2,C3,8,E5:E8). Повторяющие диапазоны обычно заданы с типом, `number[][][]` так как они являются трехмерными матрицами. В примере см. основной пример, указанный для [повторяющих параметров](#repeating-parameters).


### <a name="declaring-repeating-parameters"></a>Объявление повторяюющихся параметров
В Typescript указать, что параметр многомерный. Например, будет  `ADD(values: number[])` указывать одномерный массив, `ADD(values:number[][])` указать двумерный массив и так далее.

В JavaScript используйте `@param values {number[]}` для одномерных массивов, `@param <name> {number[][]}` для двухмерных массивов и так далее для большего размера.

Для JSON от руки убедитесь, `"repeating": true` что параметр указан как в файле JSON, а также убедитесь, что параметры помечены как `"dimensionality": matrix`.

## <a name="invocation-parameter"></a>Параметр Вызов

Каждая настраиваемая функция автоматически `invocation` передается аргументу в качестве последнего параметра ввода, даже если он явно не объявлен. Этот `invocation` параметр соответствует объекту [Вызов](/javascript/api/custom-functions-runtime/customfunctions.invocation) . Объект `Invocation` можно использовать для получения дополнительного контекста, например адреса ячейки, вызываемой на настраиваемую функцию. Чтобы получить доступ к `Invocation` объекту, необходимо `invocation` объявить его последним параметром в настраиваемой функции. 

> [!NOTE]
> Параметр `invocation` не появляется в качестве настраиваемой аргумента функции для пользователей в Excel.

В следующем примере показано, `invocation` как использовать параметр для возврата адреса ячейки, вызываемой вашей настраиваемой функцией. В этом примере используется [свойство адресов](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-address-member) `Invocation` объекта. Чтобы получить доступ к `Invocation` объекту, сначала заявите `CustomFunctions.Invocation` в качестве параметра в JSDoc. Далее заявите `@requiresAddress` в JSDoc, чтобы получить доступ к `address` свойству `Invocation` объекта. Наконец, в пределах функции извлекаем и возвращаем `address` свойство. 

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

В Excel настраиваемая `address` `Invocation` `SheetName!RelativeCellAddress` функция, вызываемая свойством объекта, возвращает абсолютный адрес после формата в ячейке, вызываемой функцией. Например, если параметр ввода расположен на листе с названием **Цены** в ячейке F6, возвращаемое значение адреса параметра будет `Prices!F6`. 

Этот `invocation` параметр также можно использовать для отправки сведений в Excel. [Дополнительные дополнительные данные см](custom-functions-web-reqs.md#make-a-streaming-function). в дополнительных данных о функциях потоковой передачи.

## <a name="detect-the-address-of-a-parameter"></a>Обнаружение адреса параметра

В сочетании с [параметром вызов](#invocation-parameter) можно использовать объект [Вызов](/javascript/api/custom-functions-runtime/customfunctions.invocation) , чтобы получить адрес настраиваемого параметра ввода функции. При вызове [параметрAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-parameteraddresses-member) свойство `Invocation` объекта позволяет функции возвращать адреса всех параметров ввода. 

Это полезно в сценариях, в которых типы входных данных могут отличаться. Адрес параметра ввода можно использовать для проверки формата номеров значения ввода. Затем формат номеров можно при необходимости скорректировать до ввода. Адрес параметра ввода также можно использовать для определения того, имеет ли значение ввода какие-либо связанные свойства, которые могут иметь отношение к последующим вычислениям. 

>[!NOTE]
> Если вы работаете с созданными вручную метаданными [JSON](custom-functions-json.md) для возврата адресов параметров вместо генератора [Yeoman для надстройок Office](../develop/yeoman-generator-overview.md), `options` `requiresParameterAddresses` объект должен иметь свойство set to `true`, `result` `dimensionality` `matrix`и объект должен иметь свойство, заданное .

Следующая настраиваемая функция состоит из трех параметров ввода, `parameterAddresses` `Invocation` извлекает свойство объекта для каждого параметра и возвращает адреса. 

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array. 
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

При запуске настраиваемой `parameterAddresses` функции, вызываемой свойством, `SheetName!RelativeCellAddress` возвращается адрес параметра после формата в ячейке, вызываемой функцией. Например, если параметр ввода расположен на листе "Затраты в ячейке D8", возвращаемое значение адреса параметра будет  .`Costs!D8` Если настраиваемая функция имеет несколько параметров и возвращается несколько адресов параметров, возвращаемые адреса будут перетекать через несколько ячеек, убывая вертикально из ячейки, вызываемой функцией. 

## <a name="next-steps"></a>Дальнейшие действия

Узнайте, как использовать [летучие значения в пользовательских функциях](custom-functions-volatile.md).

## <a name="see-also"></a>См. также

* [Получение и обработка данных с помощью пользовательских функций](custom-functions-web-reqs.md)
* [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Вручную создайте метаданные JSON для пользовательских функций](custom-functions-json.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
