---
ms.date: 04/29/2020
description: Локализация пользовательских функций Excel.
title: Локализация пользовательских функций
localization_priority: Normal
ms.openlocfilehash: 001045f82634d7e96c4d4515ccd87b5cfaf2cd1c
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275968"
---
# <a name="localize-custom-functions"></a>Локализация пользовательских функций

Вы можете локализовать как надстройку, так и имена пользовательских функций. Для этого укажите в XML-файле манифеста локализованные имена функций в файле данных JSON и сведения о языковых стандартах.

>[!IMPORTANT]
> Автоматически созданные метаданные не работают для локализации, поэтому необходимо вручную обновить файл JSON. Чтобы узнать, как это сделать, просмотрите [метаданные для пользовательских функций в Excel](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>Локализация имен функций

Чтобы локализовать пользовательские функции, создайте новый файл метаданных JSON для каждого языка. В файле JSON каждого языка создайте `name` и создайте `description` Свойства на целевом языке. Файл по умолчанию для английского языка называется **functions. JSON**. Используйте языковой стандарт в имени файла для каждого дополнительного JSON файла, например **functions – de. JSON** , чтобы определить их.

`name` `description` Они отображаются в Excel и локализуются. Тем не менее, `id` каждая функция не локализована. `id`Свойство определяет, как Excel определяет функцию как уникальную и не должна изменяться после ее задания.

В следующем коде JSON показано, как определить функцию со `id` свойством "умножить". `name`Свойство and `description` функции локализовано для немецкого языка. Каждый параметр `name` `description` также локализован для немецкого языка.

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

После создания файла JSON для каждого языка обновите XML-файл манифеста, указав значение переопределения для каждого языкового стандарта, который задает URL-адрес каждого файла метаданных JSON. Следующий XML-код манифеста показывает `en-us` языковой стандарт по умолчанию с переопределением URL-адреса файла JSON для `de-de` (Германия). Файл **functions – de. JSON** содержит локализованные имена и идентификаторы немецких функций.

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

Дополнительные сведения о процессе локализации надстройки приведены в разделе [Локализация надстроек Office](../develop/localization.md#control-localization-from-the-manifest).

## <a name="next-steps"></a>Дальнейшие действия
Сведения о [соглашениях об именовании для пользовательских функций](custom-functions-naming.md) или о том, как найти [рекомендации по обработке ошибок](custom-functions-errors.md).

## <a name="see-also"></a>См. также

* [Метаданные пользовательских функций](custom-functions-json.md)
* [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
