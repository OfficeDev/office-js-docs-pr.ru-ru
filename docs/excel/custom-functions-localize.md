---
ms.date: 11/06/2020
description: Локализовать Excel настраиваемые функции.
title: Локализация настраиваемой функции
ms.localizationpriority: medium
ms.openlocfilehash: 7219c838cfd5a6c827b74b5d04442280be7ebac7
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744507"
---
# <a name="localize-custom-functions"></a>Локализация настраиваемой функции

Можно локализовать как свои надстройки, так и настраиваемые имена функций. Для этого укажи локализованные имена функций в файле JSON функций и сведения о локализации в XML-файле манифеста.

>[!IMPORTANT]
> Автогенерированные метаданные не работают для локализации, поэтому необходимо обновить файл JSON вручную. Чтобы узнать, как это сделать, см. в инструкции по созданию метаданных [JSON для пользовательских функций](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>Локализация имен функций

Чтобы локализовать настраиваемые функции, создайте новый файл метаданных JSON для каждого языка. В каждом файле JSON языка создавайте `name` и `description` свойства на целевом языке. Файл по умолчанию для английского языка называется **functions.json**. Чтобы определить их, используйте локаль в имени файла для каждого дополнительного **JSON-файла, например functions-de.json** .

И `name` отображаются `description` в Excel и локализованы. Однако каждая `id` функция не локализована. Свойством `id` является то, Excel определяет вашу функцию как уникальную и не следует менять после ее задатки.

В следующем JSON показано, как определить функцию с свойством `id` "MULTIPLY". Свойство `name` функции `description` локализовано для немецкого языка. Каждый параметр `name` также `description` локализован для немецкого языка.

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

Создав файл JSON для каждого языка, обновите XML-файл манифеста с переопределяемым значением для каждого языка, которое указывает URL-адрес каждого файла метаданных JSON. В следующем манифесте XML показан `en-us` локальный адрес по умолчанию с URL-адресом файла JSON для `de-de` (Германия). Файл **functions-de.json содержит** локализованные имена и ids немецких функций.

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

Дополнительные сведения о процессе локализации надстройки см. в Office [надстройки](../develop/localization.md#control-localization-from-the-manifest).

## <a name="next-steps"></a>Дальнейшие действия
Узнайте о [том, как назвать условности для настраиваемой функции](custom-functions-naming.md) или найти оптимальные методы обработки [ошибок](custom-functions-errors.md).

## <a name="see-also"></a>См. также

* [Вручную создайте метаданные JSON для пользовательских функций](custom-functions-json.md)
* [Автоматическое генерирование метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
