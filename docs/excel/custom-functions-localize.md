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
# <a name="localize-custom-functions"></a><span data-ttu-id="95eed-103">Локализация пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="95eed-103">Localize custom functions</span></span>

<span data-ttu-id="95eed-104">Вы можете локализовать как надстройку, так и имена пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="95eed-104">You can localize both your add-in and your custom function names.</span></span> <span data-ttu-id="95eed-105">Для этого укажите в XML-файле манифеста локализованные имена функций в файле данных JSON и сведения о языковых стандартах.</span><span class="sxs-lookup"><span data-stu-id="95eed-105">To do so, provide localized function names in the functions' JSON file and locale information in the XML manifest file.</span></span>

>[!IMPORTANT]
> <span data-ttu-id="95eed-106">Автоматически созданные метаданные не работают для локализации, поэтому необходимо вручную обновить файл JSON.</span><span class="sxs-lookup"><span data-stu-id="95eed-106">Auto-generated metadata doesn't work for localization so you need to update the JSON file manually.</span></span> <span data-ttu-id="95eed-107">Чтобы узнать, как это сделать, просмотрите [метаданные для пользовательских функций в Excel](custom-functions-json.md)</span><span class="sxs-lookup"><span data-stu-id="95eed-107">To learn how to do this, see [Metadata for custom functions in Excel](custom-functions-json.md)</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a><span data-ttu-id="95eed-108">Локализация имен функций</span><span class="sxs-lookup"><span data-stu-id="95eed-108">Localize function names</span></span>

<span data-ttu-id="95eed-109">Чтобы локализовать пользовательские функции, создайте новый файл метаданных JSON для каждого языка.</span><span class="sxs-lookup"><span data-stu-id="95eed-109">To localize your custom functions, create a new JSON metadata file for each language.</span></span> <span data-ttu-id="95eed-110">В файле JSON каждого языка создайте `name` и создайте `description` Свойства на целевом языке.</span><span class="sxs-lookup"><span data-stu-id="95eed-110">In each language JSON file, create `name` and `description` properties in the target language.</span></span> <span data-ttu-id="95eed-111">Файл по умолчанию для английского языка называется **functions. JSON**.</span><span class="sxs-lookup"><span data-stu-id="95eed-111">The default file for English is named **functions.json**.</span></span> <span data-ttu-id="95eed-112">Используйте языковой стандарт в имени файла для каждого дополнительного JSON файла, например **functions – de. JSON** , чтобы определить их.</span><span class="sxs-lookup"><span data-stu-id="95eed-112">Use the locale in the filename for each additional JSON file, such as **functions-de.json** to help identify them.</span></span>

<span data-ttu-id="95eed-113">`name` `description` Они отображаются в Excel и локализуются.</span><span class="sxs-lookup"><span data-stu-id="95eed-113">The `name` and `description` appear in Excel and are localized.</span></span> <span data-ttu-id="95eed-114">Тем не менее, `id` каждая функция не локализована.</span><span class="sxs-lookup"><span data-stu-id="95eed-114">However, the `id` of each function isn't localized.</span></span> <span data-ttu-id="95eed-115">`id`Свойство определяет, как Excel определяет функцию как уникальную и не должна изменяться после ее задания.</span><span class="sxs-lookup"><span data-stu-id="95eed-115">The `id` property is how Excel identifies your function as unique and shouldn't be changed once it is set.</span></span>

<span data-ttu-id="95eed-116">В следующем коде JSON показано, как определить функцию со `id` свойством "умножить".</span><span class="sxs-lookup"><span data-stu-id="95eed-116">The following JSON shows how to define a function with the `id` property "MULTIPLY."</span></span> <span data-ttu-id="95eed-117">`name`Свойство and `description` функции локализовано для немецкого языка.</span><span class="sxs-lookup"><span data-stu-id="95eed-117">The `name` and `description` property of the function is localized for German.</span></span> <span data-ttu-id="95eed-118">Каждый параметр `name` `description` также локализован для немецкого языка.</span><span class="sxs-lookup"><span data-stu-id="95eed-118">Each parameter `name` and `description` is also localized for German.</span></span>

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

<span data-ttu-id="95eed-119">Сравните предыдущий JSON со следующим JSON для английского языка.</span><span class="sxs-lookup"><span data-stu-id="95eed-119">Compare the previous JSON with the following JSON for English.</span></span>

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

## <a name="localize-your-add-in"></a><span data-ttu-id="95eed-120">Локализация надстройки</span><span class="sxs-lookup"><span data-stu-id="95eed-120">Localize your add-in</span></span>

<span data-ttu-id="95eed-121">После создания файла JSON для каждого языка обновите XML-файл манифеста, указав значение переопределения для каждого языкового стандарта, который задает URL-адрес каждого файла метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="95eed-121">After creating a JSON file for each language, update your XML manifest file with an override value for each locale that specifies the URL of each JSON metadata file.</span></span> <span data-ttu-id="95eed-122">Следующий XML-код манифеста показывает `en-us` языковой стандарт по умолчанию с переопределением URL-адреса файла JSON для `de-de` (Германия).</span><span class="sxs-lookup"><span data-stu-id="95eed-122">The following manifest XML shows a default `en-us` locale with an override JSON file URL for `de-de` (Germany).</span></span> <span data-ttu-id="95eed-123">Файл **functions – de. JSON** содержит локализованные имена и идентификаторы немецких функций.</span><span class="sxs-lookup"><span data-stu-id="95eed-123">The **functions-de.json** file contains the localized German function names and ids.</span></span>

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

<span data-ttu-id="95eed-124">Дополнительные сведения о процессе локализации надстройки приведены в разделе [Локализация надстроек Office](../develop/localization.md#control-localization-from-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="95eed-124">For more information on the process of localizing an add-in, see [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).</span></span>

## <a name="next-steps"></a><span data-ttu-id="95eed-125">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="95eed-125">Next steps</span></span>
<span data-ttu-id="95eed-126">Сведения о [соглашениях об именовании для пользовательских функций](custom-functions-naming.md) или о том, как найти [рекомендации по обработке ошибок](custom-functions-errors.md).</span><span class="sxs-lookup"><span data-stu-id="95eed-126">Learn about [naming conventions for custom functions](custom-functions-naming.md) or discover [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="95eed-127">См. также</span><span class="sxs-lookup"><span data-stu-id="95eed-127">See also</span></span>

* [<span data-ttu-id="95eed-128">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="95eed-128">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="95eed-129">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="95eed-129">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="95eed-130">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="95eed-130">Create custom functions in Excel</span></span>](custom-functions-overview.md)
