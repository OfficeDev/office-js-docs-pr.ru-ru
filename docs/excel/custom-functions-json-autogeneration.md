---
title: Автоматическое генерирование метаданных JSON для пользовательских функций
description: Использование тегов JSDoc для динамического создания метаданных JSON пользовательских функций.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: da51afbcc56a86d74a9ab4edf2ebf283436196d5
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958407"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>Автоматическое генерирование метаданных JSON для пользовательских функций

Если пользовательская функция Excel написана в JavaScript или TypeScript, [теги JSDoc](https://jsdoc.app/) используются для предоставления дополнительной информации о пользовательской функции. Теги JSDoc используются при сборке для создания файла метаданных JSON. Использование тегов JSDoc позволяет не изменять файл метаданных [JSON вручную](custom-functions-json.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Добавьте тег `@customfunction` в примечаниях к коду для функции JavaScript или TypeScript, чтобы пометить ее как пользовательскую.

Типы параметров функции можно получить с помощью тега [@param](#param) в JavaScript или из раздела [Тип функции](https://www.typescriptlang.org/docs/handbook/functions.html) в TypeScript. Дополнительные сведения см. в разделах, посвященных тегу [@param](#param) и [типам](#types).

## <a name="add-a-description-to-a-function"></a>Добавление описания в функцию

Описание отображается пользователю в качестве текста справки, если ему непонятно действие пользовательской функции. Описанию не требуется какой-либо конкретный тег. Просто введите краткий текст описания в комментарии JSDoc. Обычно описание размещается в начале раздела комментариев JSDoc, но оно поддерживается независимо от места размещения.

Чтобы просмотреть примеры описаний встроенных функций, откройте Excel, перейдите на вкладку **Формулы** и нажмите кнопку **Вставить функцию**. Вы сможете просмотреть все описания функций, а также список собственных пользовательских функций.

В следующем примере фраза "Calculates the volume of a sphere." (Вычисляет объем сферы) является описанием пользовательской функции.

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```

## <a name="jsdoc-tags"></a>Теги JSDoc

Следующие теги JSDoc поддерживаются в пользовательских функциях Excel.

- [@cancelable](#cancelable)
- [@customfunction](#customfunction) *идентификатора* 
- [@helpurl URL-адрес](#helpurl) 
- [@param](#param) *имя {type}*  
- [@requiresAddress](#requiresAddress)
- [@requiresParameterAddresses](#requiresParameterAddresses)
- [@returns](#returns) *{type}*
- [@streaming](#streaming)
- [@volatile](#volatile)

---
<a id="cancelable"></a>

### <a name="cancelable"></a>@cancelable

Указывает, что пользовательская функция выполняет действие при отмене функции.

В качестве типа последнего параметра функции должно быть указано `CustomFunctions.CancelableInvocation`. Функция может назначить функцию свойству `oncanceled` , чтобы обозначить результат при отмене функции.

Если тип последнего параметра функции `CustomFunctions.CancelableInvocation`, он будет рассматриваться как `@cancelable`, даже если тег отсутствует.

Функция не может содержать одновременно теги `@cancelable` и `@streaming`.

<a id="customfunction"></a>

### <a name="customfunction"></a>@customfunction

Синтаксис: @customfunction *id* *name*

Этот тег указывает, что функция JavaScript или TypeScript является пользовательской функцией Excel. Необходимо создать метаданные для пользовательской функции.

Ниже приведен пример этого тега.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

Определяет `id` пользовательскую функцию.

- Если `id` не указан, название функции JavaScript или TypeScript преобразуется в верхний регистр, а недопустимые символы удаляются.
- `id` должен быть уникальным для всех пользовательских функций.
- Допустимые символы: A — Z, a — z, 0–9, символ подчеркивания (\_) и точка (.).

В следующем примере increment — это параметр `id` и `name` функции.

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a>name

Предоставляет отображаемый параметр `name` для пользовательской функции.

- Если имя не указано, идентификатор также используется как имя.
- Допустимые символы: буквы [буквенные символы Юникод](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), числа, точки (.) и подчеркивания (\_).
- Имя должно начинаться с буквы.
- Максимальная длина: 128 символов.

В следующем примере INC — это параметр `id` функции, а `increment` — параметр `name`.

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a>description

При вводе функции пользователям в Excel отображается описание, указывающее, что делает функция. Описанию не требуется какой-либо конкретный тег. Создайте описание для пользовательской функции, добавив в комментарии JSDoc фразу, описывающую действие функции. По умолчанию любой текст без тегов в разделе комментариев JSDoc является описанием функции.

В следующем примере фраза "A function that adds two numbers" (Функция, складывающая два числа) — это описание пользовательской функции со свойством id, имеющим значение `ADD`.

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>

### <a name="helpurl"></a>@helpurl

Синтаксис: @helpurl *url*

Предоставленный *url*-адрес отображается в Excel.

В следующем примере это `helpurl` .`www.contoso.com/weatherhelp`

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>

### <a name="param"></a>@param

#### <a name="javascript"></a>JavaScript

Синтаксис JavaScript: @param имя {type  *}*

- `{type}` указывает сведения о типе в фигурных скобках. Дополнительную информацию о типах, которые могут использоваться, см. в разделе [Типы](#types). Если тип не указан, будет использоваться тип по `any` умолчанию.
- `name` указывает параметр, к которому @param тег. Это обязательно.
- `description` предоставляет описание, которое отображается в Excel для параметра функции. Это необязательный параметр.

Чтобы обозначить параметр пользовательской функции как необязательный, поместите вокруг имени параметра квадратные скобки. Например, `@param {string} [text] Optional text`.

> [!NOTE]
> Значение по умолчанию для дополнительных параметров — `null`.

В следующем примере показана функция ADD, которая добавляет два или три числа с третьим числом в качестве необязательного параметра.

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a>TypeScript

Синтаксис TypeScript: @param *имени* 

- `name` указывает параметр, к которому @param тег. Это обязательно.
- `description` предоставляет описание, которое отображается в Excel для параметра функции. Это необязательный параметр.

Дополнительные сведения о типах параметров функций, которые могут использоваться, см. в разделе [Типы](#types).

Чтобы обозначить параметр пользовательской функции как необязательный, выполните одно из указанных ниже действий.

- Используйте необязательный параметр. Пример: `function f(text?: string)`
- Задайте для параметра значение по умолчанию. Пример: `function f(text: string = "abc")`

Подробное описание @param см. в [JSDoc](https://jsdoc.app/tags-param.html)

> [!NOTE]
> Значение по умолчанию для дополнительных параметров — `null`.

В следующем примере показана функция `add`, складывающая два числа.

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a>@requiresAddress

Указывает, что следует предоставлять адрес ячейки, в которой вычисляется функция.

Последний параметр функции должен иметь тип или `CustomFunctions.Invocation` производный тип для использования `@requiresAddress`. При вызове функции свойство `address` будет содержать адрес.

В следующем примере показано, `invocation` как использовать параметр в сочетании с `@requiresAddress` возвратом адреса ячейки, вызвавской пользовательскую функцию. [Дополнительные сведения см. в](custom-functions-parameter-options.md#invocation-parameter) параметре вызова.

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

<a id="requiresParameterAddresses"></a>

### <a name="requiresparameteraddresses"></a>@requiresParameterAddresses

Указывает, что функция должна возвращать адреса входных параметров.

Последний параметр функции должен иметь тип или `CustomFunctions.Invocation` производный тип для использования  `@requiresParameterAddresses`. Комментарий JSDoc также `@returns` должен содержать тег, указывающий, что возвращаемое значение является матрицей, например `@returns {string[][]}` или `@returns {number[][]}`. [Дополнительные сведения см](#matrix-type). в разделе "Типы матриц".

При вызове функции свойство `parameterAddresses` будет содержать адреса входных параметров.

В следующем примере показано, `invocation` как использовать параметр в сочетании с `@requiresParameterAddresses` возвратом адресов трех входных параметров. [Дополнительные сведения см](custom-functions-parameter-options.md#detect-the-address-of-a-parameter). в разделе "Определение адреса параметра".

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

<a id="returns"></a>

### <a name="returns"></a>@returns

Синтаксис: @returns {*type*}

Предоставляет тип для возвращаемого значения.

Если `{type}` не указан, будет использоваться информация о типе TypeScript. Если информация о типе отсутствует, будет использоваться тип `any`.

В следующем примере показана функция `add`, использующая тег `@returns`.

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="streaming"></a>

### <a name="streaming"></a>@streaming

Используется для обозначения того, что пользовательская функция является потоковой передачей функции.

Последний параметр имеет тип `CustomFunctions.StreamingInvocation<ResultType>`.
Функция возвращает значение `void`.

Функции потоковой передачи не возвращают значения напрямую, а вызовы с `setResult(result: ResultType)` использованием последнего параметра.

Исключения, которые возникают при потоковой передаче функций, игнорируются. `setResult()` при вызове может вернуть ошибку в качестве результата. Пример функции потоковой передачи и дополнительные сведения см. в разделе [Создание функции потоковой передачи](custom-functions-web-reqs.md#make-a-streaming-function).

Потоковые передачи функций невозможно пометить как [@volatile](#volatile).

<a id="volatile"></a>

### <a name="volatile"></a>@volatile

Переменные функции — это такие функции, чей результат не остается неизменным в каждый период времени, даже если они не содержат аргументов или их аргументы не меняются. Excel повторно проводит вычисления в ячейках, которые содержат переменные функции, вместе со всеми зависимыми функциями при каждом вычислении. По этой причине чрезмерное использование переменных функций может замедлить пересчет, поэтому используйте их умеренно.

Потоковые передачи функций не могут быть переменными.

Следующая функция является переменной и использует тег `@volatile`.

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a>Типы

Указывая тип параметра, Excel преобразует значения в этот тип, прежде чем вызывать функцию. Если указан тип `any`, преобразование выполняться не будет.

### <a name="value-types"></a>Типы значений

Одно значение может быть представлено с помощью одного из приведенных ниже типов: `boolean`, `number`, `string`.

### <a name="matrix-type"></a>Тип "матрица"

Используйте тип двумерного массива, чтобы параметр или возвращаемое значение представляли собой матрицу значений. Например, тип указывает `number[][]` матрицу `string[][]` чисел и матрицу строк.

### <a name="error-type"></a>Тип "ошибка"

Функция непотоковой передачи может указывать на ошибку, возвращая тип "Ошибка".

Функция потоковой передачи может указывать на ошибку, вызывая метод `setResult()` типа "Ошибка".

### <a name="promise"></a>Обещание

Пользовательская функция может возвращать обещание, которое предоставляет значение, когда обещание разрешается. Если обещание отклонено, пользовательская функция выдаст ошибку.

### <a name="other-types"></a>Другие типы

Любой другой тип будет рассматриваться как ошибка.

## <a name="next-steps"></a>Дальнейшие действия

Узнайте о [соглашениях именования для пользовательских функций](custom-functions-naming.md). Или же узнайте, как [локализовать свои функции](custom-functions-localize.md), для чего нужно [записать файл JSON вручную](custom-functions-json.md).

## <a name="see-also"></a>Дополнительные ресурсы

- [Создание метаданных JSON вручную для пользовательских функций](custom-functions-json.md)
- [Создание пользовательских функций в Excel](custom-functions-overview.md)
