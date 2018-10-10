---
ms.date: 09/27/2018
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel
ms.openlocfilehash: e8af13b8855d6c5e1a3b1ce99edb24445e066756
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459240"
---
# <a name="custom-functions-metadata-preview"></a>Метаданные для настраиваемых функций (предварительная версия)

При определении [настраиваемых функций](custom-functions-overview.md) в надстройке Excel необходимо включить в проект вашей надстройки файл метаданных JSON, содержащий информацию о том, что требуется Excel для того, чтобы зарегистрировать настраиваемые функции и сделать их доступными для пользователей. В этой статье описывается формат файла метаданных JSON.

Сведения о других файлах, котрые необходимо включить в проект надстройки для включения настраиваемых функций, см. в статье [Создание настраиваемых функций в Excel](custom-functions-overview.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a>Пример метаданных

В следующем примере кода показано содержимое файла метаданных JSON для надстройки, определяющей настраиваемые функции. В следующих за этим примером разделах приводится подробная информация об отдельных свойствах, представленных в данном примере JSON.

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
> Пример готового файла JSON приводится в репозитории GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## <a name="functions"></a>functions 

 Свойство `functions` представляет собой массив объектов настраиваемых функций. В следующей таблице приведены свойства каждого объекта.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Нет  |  Описание функции, которое пользователи видят в Excel. Например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта.**. |
|  `helpUrl`  |  string  |   Нет  |  URL-адрес, который предоставляет сведения о функции. (Отображается в области задач). Например, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Да | Уникальный идентификатор функции. Изменение этого идентификатора после его настройки не допускается. |
|  `name`  |  string  |  Да  |  Название функции, которое пользователи видят в Excel. В Excel название этой функции будет присоединено в качестве приставки пространством имен настраиваемой функции, указанным в XML-файле манифеста. |
|  `options`  |  object  |  Нет  |  Это свойство позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию. См. [Объект параметров](#options-object)  для получения дополнительной информации. |
|  `parameters`  |  array  |  Да  |  Массив, который определяет входные параметры для функции. См. [Массив параметров ](#parameters-array) для получения дополнительной информации. |
|  `result`  |  object  |  Да  |  Объект, который определяет тип возвращаемой функцией информации. См. [Объект результата](#result-object) для получения дополнительной информации. |

## <a name="options"></a>options

Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет данные функции. В следующей таблице описываются свойства объекта  `options`.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  Нет<br/><br/>Значение по умолчанию: `false`.  |  Если `true`, Excel будет вызывать обработчик `onCanceled` каждый раз, когда пользователь будет предпринимать действия, которые имеют эффект отмены функции, например, вручную вызывая пересчет или редактирование ячейки, на которую ссылается функция. Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным `caller`  параметром. (Не *** *** регистрируйте свои параметры в свойстве `parameters`). В теле функции необходимо назначить обработчик члену `caller.onCanceled`. Для получения дополнительной информации см.  [Отмена функции](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  boolean  |  Нет<br/><br/>Значение по умолчанию: `false`.  |  Если `true`, функция может выводить значение в ячейку несколько раз, даже если была вызвана всего единожды. Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций. Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`. (Не *** *** регистрируйте свои параметры в свойстве `parameters`). Функция должна содержать оператор `return`. Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`. Для получения дополнительной информации см. статью [Потоковые функции](custom-functions-overview.md#streaming-functions). |

## <a name="parameters"></a>parameters

Свойство `parameters`  представляет собой массив параметров объекта. В следующей таблице приводятся свойства каждого объекта.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Нет |  Описание параметра.  |
|  `dimensionality`  |  string  |  Нет  |  Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).  |
|  `name`  |  string  |  Да  |  Имя параметра. Это имя отображается в IntelliSense Excel.  |
|  `type`  |  string  |  Нет  |  Тип данных параметра. Должен представлять собой значение типа  **boolean**, **number** или **string**.  |

## <a name="result"></a>result

Объект  `results` определяет тип возвращаемой функцией информации. В следующей таблице описываются свойства объекта `result`.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Нет  |  Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив). |
|  `type`  |  string  |  Да  |  Тип данных параметра. Должен представлять собой значение типа  **boolean**, **number** или **string**.  |

## <a name="see-also"></a>См. также

* [Создание настраиваемых функций в Excel](custom-functions-overview.md)
* [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md)
* [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md)
* [Руководство по настраиваемым функциям Excel](excel-tutorial-custom-functions.md)