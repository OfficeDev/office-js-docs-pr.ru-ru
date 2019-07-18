---
ms.date: 07/10/2019
description: Узнайте о требованиях к именам пользовательских функций Excel и Избегайте распространенных ловушек именования.
title: Рекомендации по именованию пользовательских функций в Excel
localization_priority: Normal
ms.openlocfilehash: 79d0bfb069fe5abefeb6d0e88428d0728f3869e3
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771535"
---
# <a name="naming-guidelines"></a><span data-ttu-id="7469a-103">Рекомендации по именованию</span><span class="sxs-lookup"><span data-stu-id="7469a-103">Naming guidelines</span></span>

<span data-ttu-id="7469a-104">Настраиваемая функция определяется свойством **ID** и **Name** в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="7469a-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span>

- <span data-ttu-id="7469a-105">Функция `id` используется для уникальной идентификации пользовательских функций в коде JavaScript.</span><span class="sxs-lookup"><span data-stu-id="7469a-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span>
- <span data-ttu-id="7469a-106">Функция `name` используется в качестве отображаемого имени, которое отображается для пользователя в Excel.</span><span class="sxs-lookup"><span data-stu-id="7469a-106">The function `name` is used as the display name that appears to a user in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="7469a-107">Функция `name` может отличаться от функции `id`, например в целях локализации.</span><span class="sxs-lookup"><span data-stu-id="7469a-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="7469a-108">Как правило, функция `name` должна оставаться такой же, как и в `id` случае, если у них нет особой причины для их различения.</span><span class="sxs-lookup"><span data-stu-id="7469a-108">In general, a function's `name` should stay the same as the `id` if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="7469a-109">Функции `name` и `id` совместно используют некоторые общие требования:</span><span class="sxs-lookup"><span data-stu-id="7469a-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="7469a-110">Функции `id` могут использовать только буквы от A до Z, цифры от 0 до девяти, подчеркивания и точки.</span><span class="sxs-lookup"><span data-stu-id="7469a-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="7469a-111">Функция `name` может использовать любые алфавитные символы Юникода, подчеркивания и точки.</span><span class="sxs-lookup"><span data-stu-id="7469a-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="7469a-112">Обе функции `name` и `id` должны начинаться с буквы и иметь не менее трех символов.</span><span class="sxs-lookup"><span data-stu-id="7469a-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="7469a-113">Excel использует прописные буквы для встроенных имен функций (например, `SUM`).</span><span class="sxs-lookup"><span data-stu-id="7469a-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="7469a-114">Таким образом, рекомендуется использовать прописные буквы для пользовательских функций `name` и `id` в качестве рекомендаций.</span><span class="sxs-lookup"><span data-stu-id="7469a-114">Therefore, consider using uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="7469a-115">Имя функции не `name` должно совпадать с именем:</span><span class="sxs-lookup"><span data-stu-id="7469a-115">A function's `name` shouldn't be named the same as:</span></span>

- <span data-ttu-id="7469a-116">Для ячеек между a1 и XFD1048576 или ячейками между ними между R1048576C16384.</span><span class="sxs-lookup"><span data-stu-id="7469a-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="7469a-117">Любая функция макроса Excel 4,0 (например `RUN`, `ECHO`).</span><span class="sxs-lookup"><span data-stu-id="7469a-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="7469a-118">Полный список этих функций представлен в статье справочный [документ по функциям макросов Excel](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span><span class="sxs-lookup"><span data-stu-id="7469a-118">For a full list of these functions, see [this Excel Macro Functions Reference document](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="7469a-119">Конфликты имен</span><span class="sxs-lookup"><span data-stu-id="7469a-119">Naming conflicts</span></span>

<span data-ttu-id="7469a-120">Если функция `name` аналогична функции `name` в уже существующей надстройке, **#REF!**</span><span class="sxs-lookup"><span data-stu-id="7469a-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="7469a-121">в книге появится сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="7469a-121">error will appear in your workbook.</span></span>

<span data-ttu-id="7469a-122">Чтобы устранить конфликт имен, измените его `name` в надстройке и повторите функцию.</span><span class="sxs-lookup"><span data-stu-id="7469a-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="7469a-123">Вы также можете удалить надстройку с конфликтующим именем.</span><span class="sxs-lookup"><span data-stu-id="7469a-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="7469a-124">Если вы тестируете надстройку в различных средах, попробуйте использовать другое пространство имен, чтобы отличать функцию (например, `NAMESPACE_NAMEOFFUNCTION`).</span><span class="sxs-lookup"><span data-stu-id="7469a-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="7469a-125">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="7469a-125">Best practices</span></span>

- <span data-ttu-id="7469a-126">Рекомендуется добавить в функцию несколько аргументов, а не создавать несколько функций с одинаковыми или похожими именами.</span><span class="sxs-lookup"><span data-stu-id="7469a-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="7469a-127">Имена функций должны указывать на действие функции, например, `=GETZIPCODE` вместо. `ZIPCODE`</span><span class="sxs-lookup"><span data-stu-id="7469a-127">Function names should indicate the action of the function, such as `=GETZIPCODE` instead of `ZIPCODE`.</span></span>
- <span data-ttu-id="7469a-128">Избегайте неоднозначных сокращений в именах функций.</span><span class="sxs-lookup"><span data-stu-id="7469a-128">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="7469a-129">Ясности важнее, чем кратко.</span><span class="sxs-lookup"><span data-stu-id="7469a-129">Clarity is more important than brevity.</span></span> <span data-ttu-id="7469a-130">Выберите имя, `=INCREASETIME` а не `=INC`.</span><span class="sxs-lookup"><span data-stu-id="7469a-130">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="7469a-131">Согласованно используйте одни и те же команды для функций, которые выполняют похожие действия.</span><span class="sxs-lookup"><span data-stu-id="7469a-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="7469a-132">`=DELETEZIPCODE` Например, используйте `=DELETEADDRESS`и, а не `=DELETEZIPCODE` и. `=REMOVEADDRESS`</span><span class="sxs-lookup"><span data-stu-id="7469a-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>
- <span data-ttu-id="7469a-133">При указании имени потоковой функции рекомендуется добавить заметку к этому результату в описании функции или добавить `STREAM` ее в конец имени функции.</span><span class="sxs-lookup"><span data-stu-id="7469a-133">When naming a streaming function, consider adding a note to that effect in the description of the function or adding `STREAM` to the end of the function's name.</span></span>

## <a name="localizing-function-names"></a><span data-ttu-id="7469a-134">Локализация имен функций</span><span class="sxs-lookup"><span data-stu-id="7469a-134">Localizing function names</span></span>

<span data-ttu-id="7469a-135">Вы можете локализовать имена функций для разных языков с помощью отдельных файлов JSON и переопределить значения в файле манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="7469a-135">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="7469a-136">Рекомендуется избегать использования функций `id` `name` , встроенных в Excel, на другом языке, так как это может конфликтовать с локализованными функциями.</span><span class="sxs-lookup"><span data-stu-id="7469a-136">As a best practice, avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="7469a-137">Полную информацию о локализации можно найти в разделе [Localize Custom functions](custom-functions-localize.md) .</span><span class="sxs-lookup"><span data-stu-id="7469a-137">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="7469a-138">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="7469a-138">Next steps</span></span>
<span data-ttu-id="7469a-139">Ознакомьтесь с рекомендациями по [обработке ошибок](custom-functions-errors.md).</span><span class="sxs-lookup"><span data-stu-id="7469a-139">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="7469a-140">См. также</span><span class="sxs-lookup"><span data-stu-id="7469a-140">See also</span></span>

* [<span data-ttu-id="7469a-141">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="7469a-141">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="7469a-142">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="7469a-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="7469a-143">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="7469a-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
