---
ms.date: 11/06/2020
description: Локализовать Excel настраиваемые функции.
title: Локализация настраиваемой функции
localization_priority: Normal
ms.openlocfilehash: b393cbb76e4993eb77df8ddbe60247c8af74c580
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938302"
---
# <a name="localize-custom-functions"></a>Локализация настраиваемой функции

Можно локализовать как свои надстройки, так и настраиваемые имена функций. Для этого укажи локализованные имена функций в файле JSON функций и сведения о локализации в XML-файле манифеста.

>[!IMPORTANT]
> Автогенерированные метаданные не работают для локализации, поэтому необходимо обновить файл JSON вручную. Чтобы узнать, как это сделать, см. в инструкции по созданию метаданных [JSON для пользовательских функций](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>Локализация имен функций

Чтобы локализовать настраиваемые функции, создайте новый файл метаданных JSON для каждого языка. В каждом файле JSON языка создавайте `name` `description` и свойства на целевом языке. Файл по умолчанию для английского языка называется **functions.js.** Используйте локализ в имени файла для каждого дополнительного **JSON-файла,** напримерfunctions-de.jsдля их идентификации.

И `name` `description` отображаются в Excel и локализованы. Однако каждая `id` функция не локализована. Свойством является то, Excel определяет вашу функцию как уникальную и не следует менять `id` после ее задатки.

В следующем JSON показано, как определить функцию с `id` свойством "MULTIPLY". Свойство `name` `description` функции локализовано для немецкого языка. Каждый параметр `name` `description` также локализован для немецкого языка.

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

Создав файл JSON для каждого языка, обновите XML-файл манифеста с переопределяемым значением для каждого языка, которое указывает URL-адрес каждого файла метаданных JSON. В следующем манифесте XML показан локальный адрес по умолчанию с URL-адресом `en-us` файла JSON для `de-de` (Германия). В **functions-de.jsфайле** содержатся локализованные немецкие имена функций и ids.

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

Дополнительные сведения о процессе локализации надстройки см. в Office [надстройки.](../develop/localization.md#control-localization-from-the-manifest)

## <a name="next-steps"></a>Дальнейшие действия
Узнайте о [соглашениях именования](custom-functions-naming.md) пользовательских функций или откройте для себя методы обработки [ошибок.](custom-functions-errors.md)

## <a name="see-also"></a>См. также

* [Вручную создайте метаданные JSON для пользовательских функций](custom-functions-json.md)
* [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
