---
ms.date: 02/08/2019
description: Узнайте о требованиях к именам пользовательских функций Excel и Избегайте распространенных ловушек именования.
title: Рекомендации по именованию пользовательских функций в Excel (Предварительная версия)
localization_priority: Normal
ms.openlocfilehash: bdf31879fb6e750fb9dea51f66c55dbc83a2dc90
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/22/2019
ms.locfileid: "30203852"
---
# <a name="naming-guidelines"></a><span data-ttu-id="5b174-103">Рекомендации по именованию</span><span class="sxs-lookup"><span data-stu-id="5b174-103">Naming guidelines</span></span>

<span data-ttu-id="5b174-104">Настраиваемая функция определяется свойством **ID** и **Name** в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="5b174-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span> <span data-ttu-id="5b174-105">Идентификатор функции используется для уникальной идентификации пользовательских функций в коде JavaScript.</span><span class="sxs-lookup"><span data-stu-id="5b174-105">The function id is used to uniquely identify custom functions in your JavaScript code.</span></span> <span data-ttu-id="5b174-106">Имя функции используется в качестве отображаемого имени, которое отображается для пользователя в Excel.</span><span class="sxs-lookup"><span data-stu-id="5b174-106">The function name is used as the display name that appears to a user in Excel.</span></span> <span data-ttu-id="5b174-107">Имя функции может отличаться от идентификатора функции, например в целях локализации.</span><span class="sxs-lookup"><span data-stu-id="5b174-107">A function name can differ from the function ID, such as for localization purposes.</span></span> <span data-ttu-id="5b174-108">Но в общем случае он должен оставаться таким же, как и идентификатор, если нет особой причины различать их.</span><span class="sxs-lookup"><span data-stu-id="5b174-108">But in general it should stay the same as the ID if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="5b174-109">Имена функций и идентификаторы функций совместно используют некоторые общие требования:</span><span class="sxs-lookup"><span data-stu-id="5b174-109">Function names and function IDs share some common requirements:</span></span>

- <span data-ttu-id="5b174-110">Они должны использовать только буквенно-цифровые символы (включая Юникод), цифры от нуля до девяти, подчеркивания и точки.</span><span class="sxs-lookup"><span data-stu-id="5b174-110">They must only use alphanumeric characters (including Unicode), the numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="5b174-111">Они должны начинаться с буквы и иметь не менее трех символов.</span><span class="sxs-lookup"><span data-stu-id="5b174-111">They must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="5b174-112">Excel использует прописные буквы для встроенных имен функций (например, `SUM`).</span><span class="sxs-lookup"><span data-stu-id="5b174-112">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="5b174-113">Поэтому рекомендуется использовать прописные буквы для имен пользовательских функций и идентификаторов функций.</span><span class="sxs-lookup"><span data-stu-id="5b174-113">Therefore, consider using uppercase letters for your custom function names and function IDs as a best practice.</span></span>

<span data-ttu-id="5b174-114">Имена функций не должны называться одинаково:</span><span class="sxs-lookup"><span data-stu-id="5b174-114">Function names shouldn't be named the same as:</span></span>

- <span data-ttu-id="5b174-115">Для ячеек между a1 и XFD1048576 или ячейками между ними между R1048576C16384.</span><span class="sxs-lookup"><span data-stu-id="5b174-115">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="5b174-116">Любая функция макроса Excel 4,0 (например `RUN`, `ECHO`).</span><span class="sxs-lookup"><span data-stu-id="5b174-116">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="5b174-117">Полный список этих функций представлен в [этой статье](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span><span class="sxs-lookup"><span data-stu-id="5b174-117">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="5b174-118">Конфликты имен</span><span class="sxs-lookup"><span data-stu-id="5b174-118">Naming conflicts</span></span>

<span data-ttu-id="5b174-119">Если имя функции совпадает с именем функции в уже существующей надстройке, то **#REF!**</span><span class="sxs-lookup"><span data-stu-id="5b174-119">If your function name is the same as a function name in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="5b174-120">в книге появится сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="5b174-120">error will appear in your workbook.</span></span>

<span data-ttu-id="5b174-121">Чтобы устранить конфликт имен, измените имя в надстройке и повторите функцию.</span><span class="sxs-lookup"><span data-stu-id="5b174-121">To fix a name conflict, change the name in your add-in and try the function again.</span></span> <span data-ttu-id="5b174-122">Вы также можете удалить надстройку с конфликтующим именем.</span><span class="sxs-lookup"><span data-stu-id="5b174-122">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="5b174-123">Если вы тестируете надстройку в различных средах, попробуйте использовать другое пространство имен, чтобы отличать функцию (например, НАМЕСПАЦЕ_НАМЕОФФУНКТИОН).</span><span class="sxs-lookup"><span data-stu-id="5b174-123">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as NAMESPACE_NAMEOFFUNCTION).</span></span>

<span data-ttu-id="5b174-124">Кроме того, следует учитывать, как пользователи могут использовать функции в надстройке.</span><span class="sxs-lookup"><span data-stu-id="5b174-124">Also consider how you'd like people to use the functions within your add-in.</span></span> <span data-ttu-id="5b174-125">Во многих случаях имеет смысл добавить в функцию несколько аргументов вместо того, чтобы создавать несколько функций с одинаковыми или похожими именами.</span><span class="sxs-lookup"><span data-stu-id="5b174-125">In many cases, it makes sense to add multiple arguments to a function rather than create multiple functions with the same or similar names.</span></span>

## <a name="see-also"></a><span data-ttu-id="5b174-126">См. также</span><span class="sxs-lookup"><span data-stu-id="5b174-126">See also</span></span>

* [<span data-ttu-id="5b174-127">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="5b174-127">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="5b174-128">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="5b174-128">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="5b174-129">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="5b174-129">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="5b174-130">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="5b174-130">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
