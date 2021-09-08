---
ms.date: 07/08/2021
description: Использование тегов JSDoc для динамического создания метаданных JSON пользовательских функций.
title: Автоматическое генерирование метаданных JSON для пользовательских функций
localization_priority: Normal
ms.openlocfilehash: b4ae61ab46de7dadb9280e731d65715adaf64630
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936378"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>Автоматическое генерирование метаданных JSON для пользовательских функций

Если пользовательская функция Excel написана в JavaScript или TypeScript, [теги JSDoc](https://jsdoc.app/) используются для предоставления дополнительной информации о пользовательской функции. Теги JSDoc используются при сборке для создания файла метаданных JSON. Использование тегов JSDoc спасает вас от попытки вручную редактировать [файл метаданных JSON.](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Добавьте тег `@customfunction` в примечаниях к коду для функции JavaScript или TypeScript, чтобы пометить ее как пользовательскую.

Типы параметров функции можно получить с помощью тега [@param](#param) в JavaScript или из раздела [Тип функции](https://www.typescriptlang.org/docs/handbook/functions.html) в TypeScript. Дополнительные сведения см. в разделах, посвященных тегу [@param](#param) и [типам](#types).

### <a name="add-a-description-to-a-function"></a>Добавление описания в функцию

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

Следующие теги JSDoc поддерживаются в Excel пользовательских функций.

* [@cancelable](#cancelable)
* [@customfunction](#customfunction) id name
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ name description
* [@requiresAddress](#requiresAddress)
* [@requiresParameterAddresses](#requiresParameterAddresses)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
<a id="cancelable"></a>
### <a name="cancelable"></a>@cancelable

Указывает, что настраиваемая функция выполняет действие при отмене функции.

В качестве типа последнего параметра функции должно быть указано `CustomFunctions.CancelableInvocation`. Функция может назначить свойству функцию, чтобы обозначить результат `oncanceled` при отмене функции.

Если тип последнего параметра функции `CustomFunctions.CancelableInvocation`, он будет рассматриваться как `@cancelable`, даже если тег отсутствует.

Функция не может содержать одновременно теги `@cancelable` и `@streaming`.

<a id="customfunction"></a>

### <a name="customfunction"></a>@customfunction

Синтаксис: @customfunction _id_ _name_

Этот тег указывает, что функция JavaScript/TypeScript является Excel настраиваемой функцией. Необходимо создать метаданные для настраиваемой функции.

Ниже показан пример этого тега.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

Эта `id` функция определяет настраиваемую функцию.

* Если `id` не указан, название функции JavaScript или TypeScript преобразуется в верхний регистр, а недопустимые символы удаляются.
* `id` должен быть уникальным для всех пользовательских функций.
* Допустимые символы: A — Z, a — z, 0–9, символ подчеркивания (\_) и точка (.).

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

* Если имя не указано, идентификатор также используется как имя.
* Допустимые символы: буквы [буквенные символы Юникод](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), числа, точки (.) и подчеркивания (\_).
* Имя должно начинаться с буквы.
* Максимальная длина: 128 символов.

В следующем примере INC — это параметр `id` функции, а `increment` — параметр `name`.

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a>description

Описание отображается пользователям в Excel при вводе функции и указывает, какие функции она делает. Описанию не требуется какой-либо конкретный тег. Создайте описание для пользовательской функции, добавив в комментарии JSDoc фразу, описывающую действие функции. По умолчанию любой текст без тегов в разделе комментариев JSDoc является описанием функции.

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

Синтаксис: @helpurl _url_

Предоставленный _url_-адрес отображается в Excel.

В следующем примере `helpurl` это `www.contoso.com/weatherhelp` .

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

Синтаксис JavaScript: @param {type} name _description_

* `{type}` указывает сведения о типе в фигурных скобки. Дополнительную информацию о типах, которые могут использоваться, см. в разделе [Типы](#types). Если не указан тип, будет использоваться тип `any` по умолчанию.
* `name` указывает параметр, к @param тег. Это необходимо.
* `description` предоставляет описание, которое отображается в Excel для параметра функции. Это необязательно.

Чтобы обозначить параметр настраиваемой функции как необязательный, поместите квадратные скобки вокруг имени параметра. Например, `@param {string} [text] Optional text`.

> [!NOTE]
> Значение по умолчанию для дополнительных параметров — `null`.

В следующем примере показана функция ADD, которая добавляет два или три номера, а третий номер — необязательный параметр.

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

Синтаксис TypeScript: @param name _description_

* `name` указывает параметр, к @param тег. Это необходимо.
* `description` предоставляет описание, которое отображается в Excel для параметра функции. Это необязательно.

Дополнительные сведения о типах параметров функций, которые могут использоваться, см. в разделе [Типы](#types).

Чтобы обозначить параметр пользовательской функции как необязательный, выполните одно из указанных ниже действий.

* Используйте необязательный параметр. Пример: `function f(text?: string)`
* Задайте для параметра значение по умолчанию. Пример: `function f(text: string = "abc")`

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

Последний параметр функции должен быть типом или производным типом `CustomFunctions.Invocation` для `@requiresAddress` использования. При вызове функции свойство `address` будет содержать адрес.

В следующем примере показано, как использовать параметр в сочетании с возвращением адреса ячейки, вызываемой `invocation` `@requiresAddress` вашей настраиваемой функцией. Дополнительные [сведения см. в параметре Вызов.](custom-functions-parameter-options.md#invocation-parameter)

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

<a id="requiresParameterAddresses"></a>
### <a name="requiresparameteraddresses"></a>@requiresParameterAddresses

Указывает, что функция должна возвращать адреса параметров ввода. 

Последний параметр функции должен быть типом или производным типом `CustomFunctions.Invocation` для  `@requiresParameterAddresses` использования. Комментарий JSDoc также должен включать тег, указывающий, что возвращаемая величина — это `@returns` матрица, например `@returns {string[][]}` или `@returns {number[][]}` . Дополнительные [сведения см.](#matrix-type) в матричных типах. 

Когда функция называется, `parameterAddresses` свойство будет содержать адреса параметров ввода.

В следующем примере показано, как использовать параметр в сочетании с возвращением `invocation` `@requiresParameterAddresses` адресов трех параметров ввода. Дополнительные [сведения см. в](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) дополнительных сведениях Об обнаружении адреса параметра. 

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

<a id="returns"></a>
### <a name="returns"></a>@returns

Синтаксис: @returns {_type_}

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

Последний параметр — тип `CustomFunctions.StreamingInvocation<ResultType>` .
Функция `void` возвращается.

Функции потоковой передачи не возвращают значения напрямую, а звонят `setResult(result: ResultType)` с помощью последнего параметра.

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

Используйте тип двумерного массива, чтобы параметр или возвращаемое значение представляли собой матрицу значений. Например, тип указывает матрицу чисел и указывает `number[][]` `string[][]` матрицу строк.

### <a name="error-type"></a>Тип "ошибка"

Функция непотоковой передачи может указывать на ошибку, возвращая тип "Ошибка".

Функция потоковой передачи может указывать на ошибку, вызывая метод `setResult()` типа "Ошибка".

### <a name="promise"></a>Обещание

Настраиваемая функция может вернуть обещание, которое предоставляет значение при его ок. Если обещание отклоняется, то настраиваемая функция будет бросать ошибку.

### <a name="other-types"></a>Другие типы

Любой другой тип будет рассматриваться как ошибка.

## <a name="next-steps"></a>Дальнейшие действия

Узнайте о [соглашениях именования для пользовательских функций](custom-functions-naming.md). Или же узнайте, как [локализовать свои функции](custom-functions-localize.md), для чего нужно [записать файл JSON вручную](custom-functions-json.md).

## <a name="see-also"></a>Дополнительные ресурсы

* [Вручную создайте метаданные JSON для пользовательских функций](custom-functions-json.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
