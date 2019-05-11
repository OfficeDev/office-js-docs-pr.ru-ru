---
ms.date: 05/08/2019
description: Ознакомьтесь с рекомендациями по разработке пользовательских функций в Excel.
title: Рекомендации в отношении пользовательских функций
localization_priority: Normal
ms.openlocfilehash: d825f5a9f14e240ca5af3c3325cb646248d99ca9
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952105"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="d6a98-103">Рекомендации в отношении пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="d6a98-103">Custom functions best practices</span></span>

<span data-ttu-id="d6a98-104">В этой статье описаны рекомендации по разработке пользовательских функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="d6a98-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="d6a98-105">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="d6a98-105">Associating function names with JSON metadata</span></span>

<span data-ttu-id="d6a98-106">Как описано в статье [Обзор пользовательских функций](custom-functions-overview.md) проект пользовательских функций должен содержать как файл метаданных JSON, так и файл сценария (JavaScript или TypeScript) для образования готовой функции.</span><span class="sxs-lookup"><span data-stu-id="d6a98-106">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="d6a98-107">Если вы используете `yo office` метаданные JSON, можно создавать их из комментариев кода.</span><span class="sxs-lookup"><span data-stu-id="d6a98-107">If you are using `yo office` the JSON metadata can be generated from the code comments.</span></span> <span data-ttu-id="d6a98-108">В противном случае вам потребуется создать файл метаданных JSON вручную.</span><span class="sxs-lookup"><span data-stu-id="d6a98-108">Otherwise you need to build the JSON metadata file manually.</span></span>

<span data-ttu-id="d6a98-109">Чтобы функция работала должным образом, необходимо связать `id` свойство функции с реализацией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d6a98-109">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="d6a98-110">Убедитесь, что существует связь, иначе функция не будет вызываться.</span><span class="sxs-lookup"><span data-stu-id="d6a98-110">Make sure there is an association, otherwise the function will not be called.</span></span> <span data-ttu-id="d6a98-111">В приведенном ниже примере кода показано, как выполнить связь с `CustomFunctions.associate()` помощью метода.</span><span class="sxs-lookup"><span data-stu-id="d6a98-111">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="d6a98-112">Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD**.</span><span class="sxs-lookup"><span data-stu-id="d6a98-112">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="d6a98-113">В следующем JSON показаны метаданные JSON, связанные с предыдущим кодом пользовательской функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d6a98-113">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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
    },
  ]
}
```


<span data-ttu-id="d6a98-114">Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="d6a98-114">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="d6a98-115">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="d6a98-115">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="d6a98-116">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла.</span><span class="sxs-lookup"><span data-stu-id="d6a98-116">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="d6a98-117">То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.</span><span class="sxs-lookup"><span data-stu-id="d6a98-117">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

* <span data-ttu-id="d6a98-118">Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d6a98-118">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="d6a98-119">Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.</span><span class="sxs-lookup"><span data-stu-id="d6a98-119">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="d6a98-120">В файле JavaScript укажите настраиваемое сопоставление функций с помощью `CustomFunctions.associate` каждой функции.</span><span class="sxs-lookup"><span data-stu-id="d6a98-120">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="d6a98-121">В приведенном ниже примере показаны метаданные JSON, соответствующие функциям, определенным в этом примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d6a98-121">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="d6a98-122">Значения `id` свойств `name` и представлены в верхнем регистре, что является лучшим вариантом при описании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="d6a98-122">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="d6a98-123">Этот код JSON необходимо добавить только в том случае, если вы готовите собственный файл JSON вручную и не используете автоматическое создание.</span><span class="sxs-lookup"><span data-stu-id="d6a98-123">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="d6a98-124">Для получения дополнительной информации об автоформировании, ознакомьтесь со статьей [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="d6a98-124">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="d6a98-125">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="d6a98-125">Additional considerations</span></span>

<span data-ttu-id="d6a98-126">Избегайте прямого или косвенного доступа к объектной модели документов (DOM) (например, с помощью jQuery) из пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="d6a98-126">Avoid accessing the Document Object Model (DOM) directly or indirectly (for example, using jQuery) from your custom function.</span></span> <span data-ttu-id="d6a98-127">В Excel для Windows, где пользовательские функции используют [среду выполнения JavaScript](custom-functions-runtime.md), пользовательские функции не могут получить доступ к модели DOM.</span><span class="sxs-lookup"><span data-stu-id="d6a98-127">In Excel on Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="d6a98-128">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="d6a98-128">Next steps</span></span>
<span data-ttu-id="d6a98-129">Узнайте, как [выполнять веб-запросы с пользовательскими функциями](custom-functions-web-reqs.md).</span><span class="sxs-lookup"><span data-stu-id="d6a98-129">Learn how to [perform web requests with custom functions](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d6a98-130">См. также</span><span class="sxs-lookup"><span data-stu-id="d6a98-130">See also</span></span>

* [<span data-ttu-id="d6a98-131">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="d6a98-131">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="d6a98-132">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="d6a98-132">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d6a98-133">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="d6a98-133">Create custom functions in Excel</span></span>](custom-functions-overview.md)
