---
ms.date: 04/30/2019
description: Локализация пользовательских функций Excel.
title: Локализация пользовательских функций (Предварительная версия)
localization_priority: Normal
ms.openlocfilehash: 1c7fba297996c8cf050eb23b34823debf87b4e88
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527326"
---
# <a name="localize-custom-functions"></a>Локализация пользовательских функций

Чтобы пользовательские функции работали во всем мире, их следует локализовать на разных языках. Для локализации пользовательских функций необходимо указать локализованные имена функций в JSON-файле функций и предоставить сведения о языковых стандартах в XML-файле манифеста. Автоматически создаваемые метаданные не работают для локализации, поэтому необходимо вручную обновить файл JSON.

## <a name="localize-function-names"></a>Локализация имен функций

Чтобы локализовать пользовательские функции, создайте новый файл метаданных JSON для каждого языка. В файле JSON каждого языка создайте и `name` `description` создайте свойства на целевом языке. Файл по умолчанию для английского языка называется **functions. JSON**. Рекомендуется использовать языковой стандарт в имени файла для каждого дополнительного JSON файла, например **functions – de. JSON** , чтобы определить их. 

`name` Они `description` отображаются в Excel и локализуются. `id` Однако каждая функция не локализована. `id` Свойство определяет, как Excel определяет функцию как уникальную и не должна изменяться после ее задания.

В следующем коде JSON показано, как определить функцию со `id` свойством "умножить". Свойство `name` and `description` функции локализовано для немецкого языка. Каждый параметр `name` `description` также локализован для немецкого языка.

```JSON
{
    "id": "MULTIPLY",
    "name": "SUMME",
    "description": "Summe zwei Zahlen",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "eins",
            "description": "Erste Nummer",
            "dimensionality": "scalar"
        },
        {
            "name": "zwei",
            "description": "Zweite Nummer",
            "dimensionality": "scalar"
        },
    ],
}
```

Сравните предыдущий JSON со следующим JSON для английского языка.

```JSON
{
    "id": "MULTIPLY",
    "name": "Multiply",
    "description": "Multiplies two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "one",
            "description": "first number",
            "dimensionality": "scalar"
        },
        {
            "name": "two",
            "description": "second number",
            "dimensionality": "scalar"
        },
    ],
}
```

## <a name="localize-your-add-in"></a>Локализация надстройки

После создания JSON-файла для каждого языка необходимо обновить XML-файл манифеста указанным значением переопределения для каждого языкового стандарта, который задает URL-адрес каждого файла метаданных JSON. Следующий XML-код манифеста показывает `en-us` языковой стандарт по умолчанию с переОПРЕДЕЛЕНИЕМ URL-адреса файла JSON для `de-de` (Германия). Файл **functions – de. JSON** содержит локализованные имена и идентификаторы немецких функций.

```XML
<DefaultLocale>en-us</DefaultLocale>
...
<Resources>
     <bt:Urls>
        <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json"/>
          <bt:Override Locale="de-de" Value="https://localhost:3000/dist/functions-de.json" />
        </bt:url>
        
     </bt:Urls>
</Resources>
```


Дополнительные сведения о процессе локализации надстройки приведены в разделе локализация надстроек [Office](../develop/localization.md#control-localization-from-the-manifest).

## <a name="see-also"></a>См. также

* [Создание пользовательских функций в Excel](custom-functions-overview.md)
* [Метаданные пользовательских функций](custom-functions-json.md)
* [Рекомендации по пользовательским функциям](custom-functions-best-practices.md)
* [Журнал изменений пользовательских функций](custom-functions-changelog.md)
* [Руководство по настраиваемым функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
