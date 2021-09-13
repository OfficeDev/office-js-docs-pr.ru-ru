---
title: Вручную создайте метаданные JSON для пользовательских функций в Excel
description: Определите метаданные JSON для настраиваемой функции в Excel связывайте свой ID функции и свойства имен.
ms.date: 08/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8f88506cd26edf130ac5d9e06351d4fb0d711806
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151269"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>Вручную создайте метаданные JSON для пользовательских функций

Как описано [](custom-functions-overview.md) в статье обзор пользовательских функций, проект пользовательских функций должен включать как файл метаданных JSON, так и файл скрипта (JavaScript или TypeScript), чтобы зарегистрировать функцию, что делает ее доступной для использования. Настраиваемые функции регистрируются, когда пользователь запускает надстройки в первый раз и после этого доступен одному пользователю во всех книгах.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

По возможности рекомендуется использовать автогенерацию JSON вместо создания собственного JSON-файла. Автогенерация менее подвержена ошибкам пользователей, и в них уже содержатся файлы `yo office` с подмостки. Дополнительные сведения о тегах JSDoc и процессе автоматическойгенерации JSON см. в сайте [Autogenerate JSON metadata for custom functions.](custom-functions-json-autogeneration.md)

Однако проект настраиваемой функции можно сделать с нуля. Этот процесс требует:

- Напишите файл JSON.
- Убедитесь, что файл манифеста подключен к файлу JSON.
- Связать функции `id` и свойства в файле скрипта для `name` регистрации функций.

На следующем изображении объясняются различия между использованием файлов леса и `yo office` написанием JSON с нуля.

![Изображение различий между использованием Yo Office и написанием собственного JSON.](../images/custom-functions-json.png)

> [!NOTE]
> Не забудьте подключить манифест к файлу JSON, который вы создаете, через раздел в файле манифеста XML, если вы `<Resources>` не используете `yo office` генератор.

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>Авторство метаданных и подключение к манифесту

Создайте файл JSON в проекте и укажи все сведения о его функциях, таких как параметры функции. См. следующий пример [метаданных](#json-metadata-example) и ссылку на [метаданные](#metadata-reference) для полного списка свойств функций.

Убедитесь, что файл манифеста XML ссылается на файл JSON в разделе, аналогично `<Resources>` следующему примеру.

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a>Пример метаданных JSON

В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции. В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.

```json
{
  "allowErrorForDataTypeAny": true,
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE",
      "description": "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "increment",
          "description": "the number to be added each time",
          "type": "number",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": true,
        "cancelable": true
      }
    },
    {
      "id": "SECONDHIGHEST",
      "name": "SECONDHIGHEST",
      "description": "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "range",
          "description": "the input range",
          "type": "number",
          "dimensionality": "matrix"
        }
      ]
    }
  ]
}
```

> [!NOTE]
> Полный пример JSON-файла доступен в истории фиксации [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub репозитория. Поскольку проект был скорректирован для автоматического создания JSON, полный пример рукописного JSON доступен только в предыдущих версиях проекта.

## <a name="metadata-reference"></a>Ссылка на метаданные

### <a name="allowerrorfordatatypeany"></a>allowErrorForDataTypeAny

Свойство `allowErrorForDataTypeAny` — это тип данных boolean. Настройка значения позволяет `true` настраиваемой функции обрабатывать ошибки в качестве значений ввода. Все параметры с типом или могут принимать ошибки в качестве значений `any` `any[][]` ввода, `allowErrorForDataTypeAny` когда установлено `true` значение . Значение по `allowErrorForDataTypeAny` умолчанию `false` .

> [!NOTE]
> В отличие от других свойств метаданных JSON, это свойство верхнего уровня и не содержит `allowErrorForDataTypeAny` под-свойств. Пример кода кода [метаданных JSON](#json-metadata-example) см. в примере формата этого свойства.

### <a name="functions"></a>functions

Свойство `functions` представляет собой массив объектов настраиваемых функций. В таблице ниже приведены свойства каждого объекта.

| Свойство      | Тип данных | Обязательный | Описание                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | Нет       | Описание функции, которое отображается пользователям в Excel (например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).                                                            |
| `helpUrl`     | string    | Нет       | URL-адрес, по которому можно получить сведения о функции (отображается в области задач). Пример: `http://contoso.com/help/convertcelsiustofahrenheit.html`.                      |
| `id`          | string    | Да      | Уникальный идентификатор для функции. Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.                                            |
| `name`        | string    | Да      | Имя функции, которое отображается пользователям в Excel. В Excel это имя функции префиксировали в настраиваемом пространстве имен функций, указанном в XML-файле манифеста. |
| `options`     | object    | Нет       | Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию. Дополнительные сведения см. в разделе [options](#options).                                                          |
| `parameters`  | array     | Да      | Массив, который определяет входные параметры для функции. Сведения [см. в параметрах.](#parameters)                                                                             |
| `result`      | object    | Да      | Объект, который определяет тип информации, возвращаемый функцией. Дополнительные сведения см. в разделе [result](#result).                                                                 |

### <a name="options"></a>options

Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию. В таблице ниже приведены свойства объекта `options`.

| Свойство          | Тип данных | Обязательный                               | Описание |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | boolean   | Нет<br/><br/>Значение по умолчанию: `false`.  | Если это свойство имеет значение `true`, Excel будет вызывать обработчик `CancelableInvocation` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция). Отменяемые функции обычно используются только для асинхронных функций, которые возвращают один результат и требуют обработки отмены запроса на данные. Функция не может использовать как свойства, так `stream` `cancelable` и свойства. |
| `requiresAddress` | boolean   | Нет <br/><br/>Значение по умолчанию: `false`. | Если `true` ваша настраиваемая функция может получить доступ к адресу вызываемой ячейки. Свойство параметра вызов содержит адрес ячейки, вызываемой `address` вашей настраиваемой [](custom-functions-parameter-options.md#invocation-parameter) функцией. Функция не может использовать как свойства, так `stream` `requiresAddress` и свойства. |
| `requiresParameterAddresses` | boolean   | Нет <br/><br/>Значение по умолчанию: `false`. | Если ваша настраиваемая функция может получить доступ `true` к адресам входных параметров функции. Это свойство должно использоваться в сочетании с свойством объекта `dimensionality` [результатов](#result) и `dimensionality` должно быть заданной для `matrix` . Дополнительные [сведения см. в](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) дополнительных сведениях Об обнаружении адреса параметра. |
| `stream`          | boolean   | Нет<br/><br/>Значение по умолчанию: `false`.  | Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды. Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций. Функция не должна содержать оператор `return`. Вместо этого результирующее значение передается как аргумент метода обратного вызова `StreamingInvocation.setResult`. Дополнительные сведения см. в [дополнительных сведениях: Сделать функцию потоковой передачи.](custom-functions-web-reqs.md#make-a-streaming-function) |
| `volatile`        | boolean   | Нет <br/><br/>Значение по умолчанию: `false`. | Если функция пересчитывает каждый раз Excel пересчет, а не только в случае изменения зависимых значений `true` формулы. Функция не может использовать как свойства, так `stream` `volatile` и свойства. Если оба замеяны свойства и свойства, свойство `stream` `volatile` летучих свойств будет `true` проигнорировано. |

### <a name="parameters"></a>parameters

Свойство `parameters` представляет собой массив объектов параметров. В таблице ниже приведены свойства каждого объекта.

|  Свойство  |  Тип данных  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Нет |  Описание параметра. Это отображается в Excel IntelliSense.  |
|  `dimensionality`  |  string  |  Нет  |  Должно быть `scalar` либо (не массивное значение), либо `matrix` (двухмерный массив).  |
|  `name`  |  string  |  Да  |  Имя параметра. Это имя отображается в Excel в IntelliSense.  |
|  `type`  |  string  |  Нет  |  Тип данных параметра. Может быть , или , что позволяет использовать любой `boolean` `number` из `string` `any` предыдущих трех типов. Если это свойство не указано, тип данных по умолчанию `any` . |
|  `optional`  | boolean | Нет | Если присвоено значение `true`, параметр не обязателен. |
|`repeating`| boolean | Нет | Если `true` параметры заполняются из указанного массива. Обратите внимание, что все повторяющие параметры по определению считаются необязательными.  |

### <a name="result"></a>result

Объект `result` определяет тип информации, возвращаемый функцией. В таблице ниже приведены свойства объекта `result`.

| Свойство         | Тип данных | Обязательный | Описание                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | Нет       | Должно быть `scalar` либо (не массивное значение), либо `matrix` (двухмерный массив). |
| `type` | string    | Нет       | Тип данных результата. Может быть , или (что позволяет использовать любой `boolean` `number` из `string` `any` предыдущих трех типов). Если это свойство не указано, тип данных по умолчанию `any` . |

## <a name="associating-function-names-with-json-metadata"></a>Сопоставление имен функций с метаданными JSON

Для правильной работы функции необходимо связать свойство функции с реализацией `id` JavaScript. Убедитесь, что существует связь, в противном случае функция не будет зарегистрирована и не может быть Excel. В следующем примере кода показано, как сделать объединение с помощью `CustomFunctions.associate()` метода. Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD**.

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

В следующем JSON показаны метаданные JSON, связанные с предыдущим пользовательским кодом JavaScript функции.

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.

- Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.

- Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла. То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.

- Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript. Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.

- В файле JavaScript укажите настраиваемую ассоциацию функций с использованием `CustomFunctions.associate` после каждой функции.

В следующем примере показаны метаданные JSON, соответствующие функциям, определенным в предыдущем примере кода JavaScript. Значения свойства и свойства находятся в верхнем шкафу, что является наилучшей практикой при `id` `name` описании пользовательских функций. Этот JSON необходимо добавить только в том случае, если вы готовите собственный JSON-файл вручную и не используете автогенерацию. Дополнительные сведения об автогенерации см. в метаданных [Autogenerate JSON для пользовательских функций.](custom-functions-json-autogeneration.md)

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="next-steps"></a>Следующие шаги

Узнайте о [лучших методах](custom-functions-naming.md) для именования функции или узнайте, как локализовать функцию с помощью описанного ранее рукописного метода JSON. [](custom-functions-localize.md)

## <a name="see-also"></a>Дополнительные материалы

- [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
- [Параметры настраиваемой функции](custom-functions-parameter-options.md)
- [Создание пользовательских функций в Excel](custom-functions-overview.md)
