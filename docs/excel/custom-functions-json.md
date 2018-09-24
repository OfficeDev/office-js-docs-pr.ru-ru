---
ms.date: 09/20/2018
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel
ms.openlocfilehash: 815b0c6e65966867d9e5d953a40ffc705a63ee63
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062146"
---
# <a name="custom-functions-metadata"></a>Метаданные настраиваемых функций

При определении [настраиваемых функций](custom-functions-overview.md) в надстройке Excel в проект надстройки необходимо включить файл метаданных JSON, содержащий информацию о том, что требуется Excel для того, чтобы зарегистрировать настраиваемые функции и сделать их доступными для пользователей. В этой статье описывается формат файла метаданных JSON.

> [!NOTE]
> Сведения о других файлах, котрые необходимо включить в проект надстройки для включения настраиваемых функций, см. в статье [Создание настраиваемых функций в Excel](custom-functions-overview.md#learn-the-basics).

## <a name="example-metadata"></a>Пример метаданных

В следующем примере показано содержимое файла метаданных JSON для надстройки, определяющей настраиваемые функции. В следующих за этим примером разделах приводится подробная информация об отдельных свойствах, рассматриваемых в данном примере JSON.

```json
{
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ADD42ASYNC",
            "name": "ADD42ASYNC",
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ISEVEN",
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "GETDAY",
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "id": "INCREMENTVALUE",
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
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
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
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
> Пример готового файла JSON приводится в [репозитории OfficeDev/Excel-Custom-Functions GitHub](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).

## <a name="functions"></a>functions 

Свойство `functions` представляет собой массив объектов настраиваемых функций. В следующей таблице приводятся свойства каждого объекта.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Нет  |  Описание функции, отображаемое в пользовательском интерфейсе Excel. К примеру, **Преобразует градусы Цельсия в градусы Фаренгейта**. |
|  `helpUrl`  |  string  |   Нет  |  URL-адрес, позволяющий пользователю получить информацию о функции. (Отображается в области задач). Например, **http://contoso.com/help/convertcelsiustofahrenheit.html**. |
| `id`     | string | Да | Уникальный идентификатор функции. Изменение этого идентификатора после его настройки не допускается. |
|  `name`  |  string  |  Да  |  Имя функции, которая будет отображаться (добавлено в пространстве имен) в пользовательском интерфейсе Excel, когда пользователь выбирает функцию. Его совпадение с именем функции, указанным при ее определении в JavaScript, не обязательно. |
|  `options`  |  object  |  Нет  |  Это свойство позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию. См. [объект параметров](#options-object) для получения дополнительной информации. |
|  `parameters`  |  array  |  Да  |  Массив, который определяет входные параметры для функции. См. [массив параметров](#parameters-array) для получения дополнительной информации. |
|  `result`  |  object  |  Да  |  Объект, который определяет тип возвращаемой функцией информации. См. [объект результата](#result-object) для получения дополнительной информации. |

## <a name="options"></a>options

Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет данные функции. В следующей таблице описываются свойства объекта `options`.

|  Свойство  |  Тип данных  |  Обязательное  |  Description  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  boolean  |  Нет, значение по умолчанию — `false`.  |  Если `true`, Excel вызывает обработчика `onCanceled` всякий раз, когда пользователь предпринимает действие, которое имеет эффект отмены функции, например, вручную вызывая пересчет или редактирование ячейки, на которую ссылается функция. Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`. (***Не***регистрируйте этот параметр в свойстве `parameters`.) В теле функции обработчик необходимо назначить члену `caller.onCanceled`. Для получения дополнительной информации см. [Отмена функции](custom-functions-overview.md#canceling-a-function). |
|  `stream`  |  boolean  |  Нет, значение по умолчанию — `false`.  |  Если `true`, функция может выводить несколько раз в ячейку даже при вызове только один раз. Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций. Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`. (***Не***регистрируйте этот параметр в свойстве `parameters`.) Функция должна иметь выписку `return`. Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`. Для получения дополнительной информации см. статью [Потоковые функции](custom-functions-overview.md#streamed-functions). |

## <a name="parameters"></a>parameters

Свойство `parameters` представляет собой массив параметров объекта. В следующей таблице приводятся свойства каждого объекта.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `description`  |  string  |  Нет |  Описание параметра.  |
|  `dimensionality`  |  string  |  Нет  |  Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).  |
|  `name`  |  string  |  Да  |  Имя параметра. Это имя отображается в IntelliSense Excel.  |
|  `type`  |  string  |  Нет  |  Тип данных параметра. Должен представлять собой значение типа **boolean**, **number** или **string**.  |

## <a name="result"></a>result

Объект `results`, определяющий тип возвращаемой функцией информации. В следующей таблице описываются свойства объекта `result`.

|  Свойство  |  Тип данных  |  Обязательное  |  Описание  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  string  |  Нет  |  Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив). |
|  `type`  |  string  |  Да  |  Тип данных параметра. Должен представлять собой значение типа **boolean**, **number** или **string**.  |

## <a name="see-also"></a>См. также

* [Создание настраиваемых функций в Excel](custom-functions-overview.md)
* [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md)
* [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md)