---
ms.date: 05/06/2020
description: Использование тегов JSDoc для динамического создания метаданных JSON пользовательских функций.
title: Автоматическое генерирование метаданных JSON для пользовательских функций
localization_priority: Normal
ms.openlocfilehash: 97cd9a06a53019c4065c4be29e46908da766ea71
ms.sourcegitcommit: 0300165295fcbd4226aa048be2fad660892d35ea
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/06/2020
ms.locfileid: "44591132"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>Автоматическое генерирование метаданных JSON для пользовательских функций

Если пользовательская функция Excel написана в JavaScript или TypeScript, [теги JSDoc](https://jsdoc.app/) используются для предоставления дополнительной информации о пользовательской функции. Теги JSDoc используются при сборке для создания [файла метаданных JSON](custom-functions-json.md). Использование тегов JSDoc освобождает вас от необходимости редактировать файл метаданных JSON вручную.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Добавьте тег `@customfunction` в примечаниях к коду для функции JavaScript или TypeScript, чтобы пометить ее как пользовательскую.

Типы параметров функции можно получить с помощью тега [@param](#param) в JavaScript или из раздела [Тип функции](https://www.typescriptlang.org/docs/handbook/functions.html) в TypeScript. Дополнительные сведения см. в разделах, посвященных тегу [@param](#param) и [типам](#types).

### <a name="adding-a-description-to-a-function"></a>Добавление описания функции

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
Ниже приведены теги JSDoc, которые поддерживаются в пользовательских функциях Excel:
* [@cancelable](#cancelable)
* [@customfunction](#customfunction) id name
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ name description
* [@requiresAddress](#requiresAddress)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

Указывает, что настраиваемая функция выполняет действие при отмене функции.

В качестве типа последнего параметра функции должно быть указано `CustomFunctions.CancelableInvocation`. Функция может назначить функцию для `oncanceled` свойства, чтобы обозначить результат при отмене функции.

Если тип последнего параметра функции `CustomFunctions.CancelableInvocation`, он будет рассматриваться как `@cancelable`, даже если тег отсутствует.

Функция не может содержать одновременно теги `@cancelable` и `@streaming`.

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

Синтаксис: @customfunction _id_ _name_

Этот тег указывает на то, что функция JavaScript/TypeScript является пользовательской функцией Excel. Необходимо создать метаданные для пользовательской функции.

Ниже приведен пример этого тега.

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

`id`Определяет пользовательскую функцию.

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

Описание отображается для пользователей в Excel при вводе функции и указывает, что делает функция. Описанию не требуется какой-либо конкретный тег. Создайте описание для пользовательской функции, добавив в комментарии JSDoc фразу, описывающую действие функции. По умолчанию любой текст без тегов в разделе комментариев JSDoc является описанием функции.

В следующем примере фраза "A function that adds two numbers" (Функция, складывающая два числа) — это описание пользовательской функции со свойством id, имеющим значение `ADD`.

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

Синтаксис: @helpurl _url_

Предоставленный _url_-адрес отображается в Excel.

В следующем примере `helpurl` используется значение `www.contoso.com/weatherhelp` .

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a>JavaScript

Синтаксис JavaScript: @param {type} name _description_

* `{type}`Указывает сведения о типе в фигурных скобках. Дополнительную информацию о типах, которые могут использоваться, см. в разделе [Типы](#types). Если тип не указан, будет использоваться тип по умолчанию `any` .
* `name`Задает параметр, к которому применяется тег @param. Это обязательное требование.
* `description` предоставляет описание, которое отображается в Excel для параметра функции. Это необязательный параметр.

Чтобы обозначить параметр пользовательской функции как необязательный:
* Поместите имя параметра в квадратные скобки. Пример: `@param {string} [text] Optional text`.

> [!NOTE]
> Значение по умолчанию для дополнительных параметров — `null`.

В приведенном ниже примере показана функция ADD, которая складывает два или три числа с третьим числом в качестве необязательного параметра.

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

* `name`Задает параметр, к которому применяется тег @param. Это обязательное требование.
* `description` предоставляет описание, которое отображается в Excel для параметра функции. Это необязательный параметр.

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

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

Указывает, что следует предоставлять адрес ячейки, в которой вычисляется функция.

Тип последнего параметра функции должен быть `CustomFunctions.Invocation` или производной от него. При вызове функции свойство `address` будет содержать адрес.

---
### <a name="returns"></a>@returns
<a id="returns"/>

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

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

Используется для обозначения того, что пользовательская функция является потоковой передачей функции. 

Последний параметр имеет тип `CustomFunctions.StreamingInvocation<ResultType>` .
Функция возвращает значение `void` .

Функции потоковой передачи не возвращают значения напрямую, а вызывают `setResult(result: ResultType)` с помощью последнего параметра.

Исключения, которые возникают при потоковой передаче функций, игнорируются. `setResult()` при вызове может вернуть ошибку в качестве результата. Пример функции потоковой передачи и дополнительные сведения см. в разделе [Создание функции потоковой передачи](./custom-functions-web-reqs.md#make-a-streaming-function).

Потоковые передачи функций невозможно пометить как [@volatile](#volatile).

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

Переменные функции — это такие функции, чей результат не остается неизменным в каждый период времени, даже если они не содержат аргументов или их аргументы не меняются. Excel повторно проводит вычисления в ячейках, которые содержат переменные функции, вместе со всеми зависимыми функциями при каждом вычислении. По этой причине чрезмерное использование переменных функций может замедлить пересчет, поэтому используйте их умеренно.

Потоковые передачи функций не могут быть переменными.

Следующая функция является переменной и использует тег `@volatile`.

```js
/**
 * Simulates rolling a 6-sided dice.
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

Используйте тип двумерного массива, чтобы параметр или возвращаемое значение представляли собой матрицу значений. Например, тип `number[][]` указывает на матрицу чисел. `string[][]` указывает на матрицу строк.

### <a name="error-type"></a>Тип "ошибка"

Функция непотоковой передачи может указывать на ошибку, возвращая тип "Ошибка".

Функция потоковой передачи может указывать на ошибку, вызывая метод `setResult()` типа "Ошибка".

### <a name="promise"></a>Обещание

Функция может возвращать обещание, которое предоставляет значение при разрешении обещаний. Если обещание отклонено, возникает ошибка.

### <a name="other-types"></a>Другие типы

Любой другой тип будет рассматриваться как ошибка.

## <a name="next-steps"></a>Дальнейшие действия
Узнайте о [соглашениях именования для пользовательских функций](custom-functions-naming.md). Или же узнайте, как [локализовать свои функции](custom-functions-localize.md), для чего нужно [записать файл JSON вручную](custom-functions-json.md).

## <a name="see-also"></a>Дополнительные ресурсы

* [Метаданные пользовательских функций](custom-functions-json.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
