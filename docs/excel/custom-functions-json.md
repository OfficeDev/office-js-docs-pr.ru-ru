---
ms.date: 11/26/2018
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel
ms.openlocfilehash: 4bdf27173c5e912aa3eba3c8661ba45dd8b453cb
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724860"
---
# <a name="custom-functions-metadata-preview"></a>Метаданные для настраиваемых функций (предварительная версия)

При определении [настраиваемых функций](custom-functions-overview.md) в надстройке Excel в проект надстройки необходимо включить JSON-файл метаданных, содержащий информацию, необходимую Excel для регистрации настраиваемых функций и предоставления пользователям доступа к ним. В этой статье описан формат JSON-файла метаданных.

Сведения о других файлах, которые необходимо добавить в проект надстройки для включения настраиваемых функций, см. в статье [Создание пользовательских функций в Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>Пример метаданных

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
        "type": "string"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
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
      "description":  "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
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
> Пример готового JSON-файла приводится в репозитории GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## <a name="functions"></a>functions 

Свойство `functions` представляет собой массив объектов настраиваемых функций. В таблице ниже приведены свойства каждого объекта.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Нет  |  Описание функции, которое отображается пользователям в Excel (например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**). |
|  `helpUrl`  |  string  |   Нет  |  URL-адрес, по которому можно получить сведения о функции (отображается в области задач). Пример: **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Да | Уникальный идентификатор для функции. Этот идентификатор может содержать только буквы, цифры и точки, а его изменение после настройки не допускается. |
|  `name`  |  string  |  Да  |  Имя функции, которое отображается пользователям в Excel. В Excel имя этой функции будет присоединено в качестве префикса пространством имен настраиваемой функции, указанным в XML-файле манифеста. |
|  `options`  |  object  |  Нет  |  Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию. Дополнительные сведения см. в разделе [options](#options). |
|  `parameters`  |  array  |  Да  |  Массив, который определяет входные параметры для функции. Дополнительные сведения см. в разделе [parameters](#parameters). |
|  `result`  |  object  |  Да  |  Объект, который определяет тип информации, возвращаемый функцией. Дополнительные сведения см. в разделе [result](#result). |

## <a name="options"></a>options

Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию. В таблице ниже приведены свойства объекта `options`.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  Нет<br/><br/>Значение по умолчанию: `false`.  |  Если это свойство имеет значение `true`, Excel будет вызывать обработчик `onCanceled` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция). Если вы используете это свойство, Excel вызовет функцию JavaScript с дополнительным параметром `caller`. (***Не*** регистрируйте этот параметр в свойстве `parameters`.) В тексте функции обработчик необходимо назначить элементу `caller.onCanceled`. Дополнительные сведения см. в разделе [Отмена функции](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  boolean  |  Нет<br/><br/>Значение по умолчанию: `false`.  |  Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды. Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций. Если вы используете это свойство, Excel вызовет функцию JavaScript с дополнительным параметром `caller`. (***Не*** регистрируйте этот параметр в свойстве `parameters`.) Функция не должна содержать оператор `return`. Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`. Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-overview.md#streaming-functions). |
|  `volatile`  | boolean | Нет <br/><br/>Значение по умолчанию: `false`. | <br /><br /> Если присвоено значение `true`, функция пересчитывается при каждом выполнении пересчета в Excel, а не только при изменении зависимых значений формулы. Функция не может быть одновременно потоковой и переменной. Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться. |

## <a name="parameters"></a>parameters

Свойство `parameters` представляет собой массив объектов параметров. В таблице ниже приведены свойства каждого объекта.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Нет |  Описание параметра. Отображается в IntelliSense Excel.  |
|  `dimensionality`  |  string  |  Нет  |  Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).  |
|  `name`  |  string  |  Да  |  Имя параметра. Оно отображается в IntelliSense Excel.  |
|  `type`  |  string  |  Нет  |  Тип данных параметра. Может иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов. Если это свойство не задано, по умолчанию устанавливается тип данных **any**. |
|  `optional`  | boolean | Нет | Если присвоено значение `true`, параметр не обязателен. |

>[!NOTE]
> Если свойство `type` необязательного параметра не указано или равно `any`, вы можете заметить проблемы, например ошибки линтинга в интегрированной среде разработки (IDE) и отсутствие необязательных параметров при вводе функции в ячейке Excel. Это планируется изменить в декабре 2018 г.

## <a name="result"></a>result

Объект `result` определяет тип информации, возвращаемый функцией. В таблице ниже приведены свойства объекта `result`.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Нет  |  Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив). |
|  `type`  |  string  |  Да  |  Тип данных параметра. Должен иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов. |

## <a name="see-also"></a>См. также

* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
* [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md)
* [Руководство по настраиваемым функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
