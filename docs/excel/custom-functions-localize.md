---
ms.date: 05/03/2019
description: Локализация пользовательских функций Excel.
title: Локализация пользовательских функций
localization_priority: Normal
ms.openlocfilehash: 5dbe2f78f1d24c3d8c8214f4e604e66f097adba3
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628034"
---
# <a name="localize-custom-functions"></a>Локализация пользовательских функций

Вы можете локализовать как надстройку, так и имена пользовательских функций. Необходимо указать локализованные имена функций в JSON-файле функций и предоставить сведения о языковых стандартах в XML-файле манифеста.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!IMPORTANT]
> Автоматически создаваемые метаданные не работают для локализации, поэтому необходимо вручную обновить файл JSON.

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

## <a name="next-steps"></a>Дальнейшие действия
Сведения о [соглашениях об именовании для пользовательских функций](custom-functions-naming.md) или о том, как найти [рекомендации по обработке ошибок](custom-functions-errors.md).

## <a name="see-also"></a>См. также

* [Метаданные пользовательских функций](custom-functions-json.md)
* [Автоматически создавать метаданные JSON для пользовательских функций](custom-functions-json-autogeneration.md)
* [Рекомендации по пользовательским функциям](custom-functions-best-practices.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
