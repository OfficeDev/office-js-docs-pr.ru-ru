---
ms.date: 11/06/2020
description: Узнайте о требованиях к именам пользовательских функций Excel и Избегайте распространенных ловушек именования.
title: Рекомендации по именованию пользовательских функций в Excel
localization_priority: Normal
ms.openlocfilehash: eefd703c63311934435657bf9e6159662f908a95
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071615"
---
# <a name="custom-functions-naming-guidelines"></a><span data-ttu-id="1b49f-103">Рекомендации по именованию пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="1b49f-103">Custom functions naming guidelines</span></span>

<span data-ttu-id="1b49f-104">Пользовательская функция определяется `id` `name` свойством и в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="1b49f-104">A custom function is identified by an `id` and `name` property in the JSON metadata file.</span></span>

- <span data-ttu-id="1b49f-105">Функция `id` используется для уникальной идентификации пользовательских функций в коде JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1b49f-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span>
- <span data-ttu-id="1b49f-106">Функция `name` используется в качестве отображаемого имени, которое отображается для пользователя в Excel.</span><span class="sxs-lookup"><span data-stu-id="1b49f-106">The function `name` is used as the display name that appears to a user in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="1b49f-107">Функция `name` может отличаться от функции `id` , например в целях локализации.</span><span class="sxs-lookup"><span data-stu-id="1b49f-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="1b49f-108">В общем случае функция `name` должна оставаться такой же, как и в `id` случае, если нет причин различать их.</span><span class="sxs-lookup"><span data-stu-id="1b49f-108">In general, a function's `name` should stay the same as the `id` if there is no reason for them to differ.</span></span>

<span data-ttu-id="1b49f-109">Функции `name` и `id` совместно используют некоторые общие требования:</span><span class="sxs-lookup"><span data-stu-id="1b49f-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="1b49f-110">Функции `id` могут использовать только буквы от A до Z, цифры от 0 до девяти, подчеркивания и точки.</span><span class="sxs-lookup"><span data-stu-id="1b49f-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="1b49f-111">Функция `name` может использовать любые алфавитные символы Юникода, подчеркивания и точки.</span><span class="sxs-lookup"><span data-stu-id="1b49f-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="1b49f-112">Обе функции `name` и `id` должны начинаться с буквы и иметь не менее трех символов.</span><span class="sxs-lookup"><span data-stu-id="1b49f-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="1b49f-113">Excel использует прописные буквы для встроенных имен функций (например, `SUM` ).</span><span class="sxs-lookup"><span data-stu-id="1b49f-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="1b49f-114">Используйте прописные буквы для пользовательских функций `name` и `id` в качестве рекомендаций.</span><span class="sxs-lookup"><span data-stu-id="1b49f-114">Use uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="1b49f-115">Функция `name` не должна быть такой же, как:</span><span class="sxs-lookup"><span data-stu-id="1b49f-115">A function's `name` shouldn't be the same as:</span></span>

- <span data-ttu-id="1b49f-116">Для ячеек между a1 и XFD1048576 или ячейками между ними между R1048576C16384.</span><span class="sxs-lookup"><span data-stu-id="1b49f-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="1b49f-117">Любая функция макроса Excel 4,0 (например `RUN` , `ECHO` ).</span><span class="sxs-lookup"><span data-stu-id="1b49f-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="1b49f-118">Полный список этих функций представлен в статье [справочный документ по функциям макросов Excel](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span><span class="sxs-lookup"><span data-stu-id="1b49f-118">For a full list of these functions, see [this Excel Macro Functions Reference document](https://d13ot9o61jdzpp.cloudfront.net/files/Excel%204.0%20Macro%20Functions%20Reference.pdf).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="1b49f-119">Конфликты имен</span><span class="sxs-lookup"><span data-stu-id="1b49f-119">Naming conflicts</span></span>

<span data-ttu-id="1b49f-120">Если функция аналогична `name` функции `name` в уже существующей надстройке, **#REF!**</span><span class="sxs-lookup"><span data-stu-id="1b49f-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="1b49f-121">в книге появится сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="1b49f-121">error will appear in your workbook.</span></span>

<span data-ttu-id="1b49f-122">Чтобы устранить конфликт имен, измените его `name` в надстройке и повторите функцию.</span><span class="sxs-lookup"><span data-stu-id="1b49f-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="1b49f-123">Вы также можете удалить надстройку с конфликтующим именем.</span><span class="sxs-lookup"><span data-stu-id="1b49f-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="1b49f-124">Если вы тестируете надстройку в различных средах, попробуйте использовать другое пространство имен, чтобы отличать функцию (например, `NAMESPACE_NAMEOFFUNCTION` ).</span><span class="sxs-lookup"><span data-stu-id="1b49f-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="1b49f-125">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="1b49f-125">Best practices</span></span>

- <span data-ttu-id="1b49f-126">Рекомендуется добавить в функцию несколько аргументов, а не создавать несколько функций с одинаковыми или похожими именами.</span><span class="sxs-lookup"><span data-stu-id="1b49f-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="1b49f-127">Избегайте неоднозначных сокращений в именах функций.</span><span class="sxs-lookup"><span data-stu-id="1b49f-127">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="1b49f-128">Ясности важнее, чем кратко.</span><span class="sxs-lookup"><span data-stu-id="1b49f-128">Clarity is more important than brevity.</span></span> <span data-ttu-id="1b49f-129">Выберите имя, `=INCREASETIME` а не `=INC` .</span><span class="sxs-lookup"><span data-stu-id="1b49f-129">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="1b49f-130">Имена функций должны указывать на действие функции, например = ЖЕТЗИПКОДЕ, а не индекс.</span><span class="sxs-lookup"><span data-stu-id="1b49f-130">Function names should indicate the action of the function, such as =GETZIPCODE instead of ZIPCODE.</span></span>
- <span data-ttu-id="1b49f-131">Согласованно используйте одни и те же команды для функций, которые выполняют похожие действия.</span><span class="sxs-lookup"><span data-stu-id="1b49f-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="1b49f-132">Например, используйте `=DELETEZIPCODE` и `=DELETEADDRESS` , а не `=DELETEZIPCODE` и `=REMOVEADDRESS` .</span><span class="sxs-lookup"><span data-stu-id="1b49f-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>
- <span data-ttu-id="1b49f-133">При указании имени потоковой функции рекомендуется добавить заметку к этому результату в описании функции или добавить ее `STREAM` в конец имени функции.</span><span class="sxs-lookup"><span data-stu-id="1b49f-133">When naming a streaming function, consider adding a note to that effect in the description of the function or adding `STREAM` to the end of the function's name.</span></span>

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="localizing-function-names"></a><span data-ttu-id="1b49f-134">Локализация имен функций</span><span class="sxs-lookup"><span data-stu-id="1b49f-134">Localizing function names</span></span>

<span data-ttu-id="1b49f-135">Вы можете локализовать имена функций для разных языков с помощью отдельных файлов JSON и переопределить значения в файле манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="1b49f-135">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="1b49f-136">Не выполняйте функции с `id` помощью `name` встроенной функции Excel на другом языке, так как это может конфликтовать с локализованными функциями.</span><span class="sxs-lookup"><span data-stu-id="1b49f-136">Avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="1b49f-137">Полную информацию о локализации можно найти в разделе [Localize Custom functions](custom-functions-localize.md) .</span><span class="sxs-lookup"><span data-stu-id="1b49f-137">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="1b49f-138">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="1b49f-138">Next steps</span></span>
<span data-ttu-id="1b49f-139">Ознакомьтесь с рекомендациями по [обработке ошибок](custom-functions-errors.md).</span><span class="sxs-lookup"><span data-stu-id="1b49f-139">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1b49f-140">См. также</span><span class="sxs-lookup"><span data-stu-id="1b49f-140">See also</span></span>

* [<span data-ttu-id="1b49f-141">Создание метаданных JSON для пользовательских функций вручную</span><span class="sxs-lookup"><span data-stu-id="1b49f-141">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="1b49f-142">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="1b49f-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
