---
title: Параметры пользовательских функций Excel
description: Узнайте, как использовать различные параметры в пользовательских функциях, например диапазоны Excel, необязательные параметры, контекст вызова и т. д.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: de86afc60d7d0b81820bd742e989e0ee7dd6970c
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958575"
---
# <a name="custom-functions-parameter-options"></a>Параметры настраиваемых функций

Настраиваемые функции можно настроить с множеством различных параметров.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>Необязательные параметры

Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках. В следующем примере функция добавления может при необходимости добавить третье число. Эта функция отображается как в `=CONTOSO.ADD(first, second, [third])` Excel.

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
> Если для необязательного параметра не указано значение, Excel присваивает ему значение `null`. Это означает, что параметры, инициализированные по умолчанию в TypeScript, не будут работать должным образом. Не используйте синтаксис, `function add(first:number, second:number, third=0):number` так как он не будет инициализироваться до `third` 0. Вместо этого используйте синтаксис TypeScript, как показано в предыдущем примере.

При указании функции, содержащего один или несколько необязательных параметров, укажите, что происходит, если необязательные параметры имеют значение NULL. В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`. Если параметр `zipCode` имеет значение NULL, по умолчанию устанавливается значение `98052`. Если параметр `dayOfWeek` имеет значение NULL, он имеет значение "Среда".

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

Пользовательская функция может принимать диапазон данных ячеек в качестве входного параметра. Функция также может возвращать диапазон данных. Excel передает диапазон данных ячеек в виде двумерного массива.

Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel. Приведенная ниже `values`функция принимает параметр, а синтаксис JSDOC `number[][]` `dimensionality` `matrix` задает для свойства параметра значение в метаданных JSON для этой функции.

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
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

Повторяющийся параметр позволяет пользователю ввести ряд необязательных аргументов в функцию. При вызове функции значения предоставляются в массиве для параметра. Если имя параметра заканчивается числом, номер каждого аргумента будет увеличиваться постепенно, `ADD(number1, [number2], [number3],…)`например . Это соответствует соглашению, используемом для встроенных функций Excel.

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

![Пользовательская функция ADD, введенная в ячейку листа Excel](../images/operands.png)

### <a name="repeating-single-value-parameter"></a>Повторяющийся параметр с одним значением

Повторяющийся параметр с одним значением позволяет передавать несколько отдельных значений. Например, пользователь может ввести ADD(1,B2,3). В следующем примере показано, как объявить один параметр значения.

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

### <a name="single-range-parameter"></a>Параметр с одним диапазоном

Параметр с одним диапазоном технически не является повторяющимся параметром, но включается здесь, так как объявление очень похоже на повторяющиеся параметры. Пользователю будет показано, как ADD(A2:B3), где один диапазон передается из Excel. В следующем примере показано, как объявить один параметр диапазона.

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

Параметр повторяющегося диапазона позволяет передавать несколько диапазонов или чисел. Например, пользователь может ввести ADD(5,B2,C3,8,E5:E8). Повторяющиеся диапазоны обычно заданы с типом, `number[][][]` так как они являются трехмерными матрицами. Пример см. в основном примере, указанном для [повторяющихся параметров](#repeating-parameters).

### <a name="declaring-repeating-parameters"></a>Объявление повторяющихся параметров

В Typescript укажите, что параметр является многомерным. Например, можно указать  `ADD(values: number[])` одномерный массив, `ADD(values:number[][])` указать двумерный массив и т. д.

В JavaScript используйте одномерные `@param values {number[]}` массивы, `@param <name> {number[][]}` для двумерных массивов и т. д. для других измерений.

Для JSON, созданных вручную, убедитесь, `"repeating": true` что параметр указан как в файле JSON, а также убедитесь, что параметры помечены как `"dimensionality": matrix`.

## <a name="invocation-parameter"></a>Параметр вызова

Каждой пользовательской функции автоматически передается аргумент `invocation` в качестве последнего входного параметра, даже если он не объявлен явным образом. Этот `invocation` параметр соответствует [объекту Вызова](/javascript/api/custom-functions-runtime/customfunctions.invocation) . Объект `Invocation` можно использовать для получения дополнительного контекста, например адреса ячейки, вызвавской пользовательскую функцию. Чтобы получить доступ к `Invocation` объекту, необходимо `invocation` объявить его последним параметром в пользовательской функции.

> [!NOTE]
> Параметр `invocation` не отображается как настраиваемый аргумент функции для пользователей в Excel.

В следующем примере показано, как использовать параметр `invocation` для возврата адреса ячейки, вызвавской пользовательскую функцию. В этом примере [используется свойство адреса](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-address-member) `Invocation` объекта. Чтобы получить доступ к `Invocation` объекту, сначала объявите `CustomFunctions.Invocation` его в качестве параметра в JSDoc. Затем объявите `@requiresAddress` в JSDoc, чтобы получить доступ к `address` свойству `Invocation` объекта. Наконец, в функции извлеките и возвращаете `address` свойство.

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
  const address = invocation.address;
  return address;
}
```

В Excel пользовательская функция, вызывающее `address` `Invocation` свойство объекта, `SheetName!RelativeCellAddress` возвращает абсолютный адрес, следующий за форматом в ячейке, вызвавской функцию. Например, если входной параметр находится на листе с именем **Prices** в ячейке F6, возвращается значение адреса параметра `Prices!F6`.

Этот `invocation` параметр также можно использовать для отправки сведений в Excel. [Дополнительные сведения см. в](custom-functions-web-reqs.md#make-a-streaming-function) статье "Создание функции потоковой передачи".

## <a name="detect-the-address-of-a-parameter"></a>Определение адреса параметра

В сочетании с [параметром](#invocation-parameter) вызова можно использовать объект [Вызова](/javascript/api/custom-functions-runtime/customfunctions.invocation) для получения адреса входного параметра пользовательской функции. При вызове свойство [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-parameteraddresses-member) `Invocation` объекта позволяет функции возвращать адреса всех входных параметров.

Это полезно в сценариях, где типы входных данных могут различаться. Адрес входного параметра можно использовать для проверки числового формата входного значения. При необходимости формат числа можно изменить перед входными данными. Адрес входного параметра также можно использовать для определения того, имеет ли входное значение какие-либо связанные свойства, которые могут быть связаны с последующими вычислениями.

>[!NOTE]
> Если вы работаете с созданными вручную метаданными [JSON](custom-functions-json.md) для возврата адресов параметров вместо генератора [Yeoman](../develop/yeoman-generator-overview.md) для надстроек Office, `options` `requiresParameterAddresses` `true`то для объекта должно быть задано значение свойства, `result` `dimensionality` а для объекта должно быть задано значение .`matrix`

Приведенная ниже пользовательская функция принимает три входных параметра, `parameterAddresses` `Invocation` извлекает свойство объекта для каждого параметра и возвращает адреса.

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
  const addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

При выполнении пользовательской функции `parameterAddresses` , вызываемой свойством, `SheetName!RelativeCellAddress` адрес параметра возвращается после формата в ячейке, вызвавской функцию. Например, если входной параметр находится на листе с именем **Costs** в ячейке D8, возвращается значение адреса параметра `Costs!D8`. Если пользовательская функция имеет несколько параметров и возвращается несколько адресов параметров, возвращаемые адреса будут перелиты по нескольким ячейкам по вертикали из ячейки, вызвавской функцию.

## <a name="next-steps"></a>Дальнейшие действия

Узнайте, как использовать [переменные значения в пользовательских функциях](custom-functions-volatile.md).

## <a name="see-also"></a>Дополнительные ресурсы

- [Получение и обработка данных с помощью пользовательских функций](custom-functions-web-reqs.md)
- [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
- [Создание метаданных JSON вручную для пользовательских функций](custom-functions-json.md)
- [Создание пользовательских функций в Excel](custom-functions-overview.md)
- [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
