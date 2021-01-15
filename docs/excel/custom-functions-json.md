---
ms.date: 12/22/2020
description: Определите метаданные JSON для пользовательских функций в Excel и привяжйте к свойствам имени и ИД функции.
title: Создание метаданных JSON вручную для пользовательских функций в Excel
localization_priority: Normal
ms.openlocfilehash: 80a71c640caacbd865b0dd253f03258a64c9b1bf
ms.sourcegitcommit: 48b9c3b63668b2a53ce73f92ce124ca07c5ca68c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2020
ms.locfileid: "49735552"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a>Создание метаданных JSON для пользовательских функций вручную

Как описано [](custom-functions-overview.md) в статье с обзором пользовательских функций, проект пользовательских функций должен включать файл метаданных JSON и файл скрипта (JavaScript или TypeScript) для регистрации функции, что делает ее доступной для использования. Пользовательские функции регистрируются, когда пользователь запускает надстройки в первый раз, а затем они доступны одному пользователю во всех книгах.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Рекомендуется по возможности использовать автоматическое создание JSON вместо создания собственного JSON-файла. Автогенерация менее подвержена ошибкам пользователя, и в них уже есть `yo office` скафолдолды. Дополнительные сведения о тегах JSDoc и процессе автоматического создания JSON см. в автогенерации метаданных JSON для [пользовательских функций.](custom-functions-json-autogeneration.md)

Однако проект пользовательских функций можно сделать с нуля. Этот процесс требует:

- Напишите файл JSON.
- Убедитесь, что файл манифеста подключен к JSON-файлу.
- Связывайте функции `id` и `name` свойства в файле скрипта, чтобы зарегистрировать функции.

На следующем рисунке поясняются различия между использованием файлов скаффолда `yo office` и написанием JSON с нуля.

![Изображение различий между использованием Yo Office и написанием собственного JSON](../images/custom-functions-json.png)

> [!NOTE]
> Не забудьте подключить манифест к JSON-файлу, если генератор не используется, с помощью раздела в `<Resources>` XML-файле `yo office` манифеста.

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a>Авторство метаданных и подключение к манифесту

Создайте JSON-файл в своем проекте и предокабъем все сведения о функциях в нем, например параметры функции. Полный [список свойств функции](#json-metadata-example) см. в следующем примере метаданных и справочнике по метаданным. [](#metadata-reference)

Убедитесь, что XML-файл манифеста ссылается на файл JSON в разделе, как в `<Resources>` следующем примере.

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
> Полный пример JSON-файла доступен в истории фиксации репозитория [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub. Так как проект был настроен для автоматического создания JSON, полный пример рукописного JSON доступен только в предыдущих версиях проекта.

## <a name="metadata-reference"></a>Справочник по метаданным

### <a name="functions"></a>functions

Свойство `functions` представляет собой массив объектов настраиваемых функций. В таблице ниже приведены свойства каждого объекта.

| Свойство      | Тип данных | Обязательный | Описание                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | string    | Нет       | Описание функции, которое отображается пользователям в Excel (например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).                                                            |
| `helpUrl`     | string    | Нет       | URL-адрес, по которому можно получить сведения о функции (отображается в области задач). Пример: `http://contoso.com/help/convertcelsiustofahrenheit.html`.                      |
| `id`          | string    | Да      | Уникальный идентификатор для функции. Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.                                            |
| `name`        | string    | Да      | Имя функции, которое отображается пользователям в Excel. В Excel это имя функции имеет префикс пространства имен пользовательских функций, указанного в XML-файле манифеста. |
| `options`     | object    | Нет       | Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию. Дополнительные сведения см. в разделе [options](#options).                                                          |
| `parameters`  | array     | Да      | Массив, который определяет входные параметры для функции. Подробные [сведения см. в](#parameters) параметрах.                                                                             |
| `result`      | object    | Да      | Объект, который определяет тип информации, возвращаемый функцией. Дополнительные сведения см. в разделе [result](#result).                                                                 |

### <a name="options"></a>options

Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию. В таблице ниже приведены свойства объекта `options`.

| Свойство          | Тип данных | Обязательный                               | Описание |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | boolean   | Нет<br/><br/>Значение по умолчанию: `false`.  | Если это свойство имеет значение `true`, Excel будет вызывать обработчик `CancelableInvocation` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция). Отменяемые функции обычно используются только для асинхронных функций, которые возвращают один результат и должны обрабатывать отмену запроса данных. Функция не может использовать свойства `stream` `cancelable` и свойства. |
| `requiresAddress` | boolean   | Нет <br/><br/>Значение по умолчанию: `false`. | Если `true` пользовательская функция может получить доступ к адресу вызываемой ячейки. Свойство параметра вызова содержит адрес ячейки, которая `address` вызывает пользовательскую [](custom-functions-parameter-options.md#invocation-parameter) функцию. Функция не может использовать свойства `stream` `requiresAddress` и свойства. |
| `requiresParameterAddresses` | boolean   | Нет <br/><br/>Значение по умолчанию: `false`. | Если пользовательская функция может получить доступ к адресам входных `true` параметров функции. Это свойство должно использоваться в сочетании со свойством объекта результата и должно `dimensionality` быть [](#result) `dimensionality` установлено в качестве `matrix` . Дополнительные [сведения см.](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) в под вопросе "Обнаружение адреса параметра". |
| `stream`          | boolean   | Нет<br/><br/>Значение по умолчанию: `false`.  | Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды. Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций. Функция не должна содержать оператор `return`. Вместо этого результирующее значение передается как аргумент метода обратного вызова `StreamingInvocation.setResult`. Дополнительные сведения см. в [функции "Make a streaming function" (Функция потоковой передачи).](custom-functions-web-reqs.md#make-a-streaming-function) |
| `volatile`        | boolean   | Нет <br/><br/>Значение по умолчанию: `false`. | Если функция пересчитывает каждый раз пересчет Excel, а не только при изменениях зависимых `true` значений формулы. Функция не может использовать свойства `stream` `volatile` и свойства. Если оба `stream` свойства `volatile` и за `true` установлены, переменное свойство будет игнорироваться. |

### <a name="parameters"></a>parameters

Свойство `parameters` представляет собой массив объектов параметров. В таблице ниже приведены свойства каждого объекта.

|  Свойство  |  Тип данных  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Нет |  Описание параметра. Это отображается в IntelliSense Excel.  |
|  `dimensionality`  |  string  |  Нет  |  Должен быть `scalar` либо (не массивное значение), либо `matrix` (двумерный массив).  |
|  `name`  |  string  |  Да  |  Имя параметра. Это имя отображается в IntelliSense Excel.  |
|  `type`  |  string  |  Нет  |  Тип данных параметра. Может быть , , , или , , который позволяет использовать `boolean` `number` любой из `string` `any` предыдущих трех типов. Если это свойство не указано, по умолчанию тип данных имеет значение `any` . |
|  `optional`  | boolean | Нет | Если присвоено значение `true`, параметр не обязателен. |
|`repeating`| boolean | Нет | Если `true` параметры заполняются из указанного массива. Обратите внимание, что все повторяющие параметры считаются необязательными по определению.  |

### <a name="result"></a>result

Объект `result` определяет тип информации, возвращаемый функцией. В таблице ниже приведены свойства объекта `result`.

| Свойство         | Тип данных | Обязательный | Описание                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | string    | Нет       | Должен быть `scalar` либо (не массивное значение), либо `matrix` (двумерный массив). |
| `type` | string    | Нет       | Тип данных результата. Может быть , , или (который позволяет использовать любой `boolean` `number` из `string` `any` предыдущих трех типов). Если это свойство не указано, по умолчанию тип данных имеет значение `any` . |

## <a name="associating-function-names-with-json-metadata"></a>Сопоставление имен функций с метаданными JSON

Для правильной работы функции необходимо связать ее свойство с реализацией `id` JavaScript. Убедитесь, что существует связь, в противном случае функция не будет зарегистрирована и не будет использоваться в Excel. В следующем примере кода показано, как сделать связь с помощью `CustomFunctions.associate()` метода. Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD**.

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

В следующем коде JSON показаны метаданные JSON, связанные с предыдущим кодом JavaScript пользовательской функции.

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

- В файле JavaScript укажите настраиваемую связь функции после `CustomFunctions.associate` каждой функции.

В следующем примере показаны метаданные JSON, соответствующие функциям, определенным в предыдущем примере кода JavaScript. Значения `id` свойств и их значения заголовочные, что лучше всего при `name` описании пользовательских функций. Этот JSON необходимо добавить только в том случае, если вы подготавливаете собственный JSON-файл вручную и не используете автогенерацию. Дополнительные сведения об автоматическомгенерации см. в сведениях о метаданных [autogenerate JSON для пользовательских функций.](custom-functions-json-autogeneration.md)

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
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

## <a name="next-steps"></a>Дальнейшие действия

Узнайте, [как назвать функцию](custom-functions-naming.md) или локализовать ее с помощью ранее описанного метода JSON. [](custom-functions-localize.md)

## <a name="see-also"></a>См. также

- [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
- [Параметры пользовательских функций](custom-functions-parameter-options.md)
- [Создание пользовательских функций в Excel](custom-functions-overview.md)
