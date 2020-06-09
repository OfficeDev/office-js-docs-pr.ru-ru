---
title: Получение и Настройка категорий
description: Как управлять категориями для почтового ящика и элемента
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d4589571de47218741308c01caec0166d72919d8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608980"
---
# <a name="get-and-set-categories"></a><span data-ttu-id="5971f-103">Получение и Настройка категорий</span><span class="sxs-lookup"><span data-stu-id="5971f-103">Get and set categories</span></span>

<span data-ttu-id="5971f-104">В Outlook пользователь может применять категории к сообщениям и встречам в виде средств Организации их данных почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="5971f-104">In Outlook, a user can apply categories to messages and appointments as a means of organizing their mailbox data.</span></span> <span data-ttu-id="5971f-105">Пользователь определяет главный список категорий для своего почтового ящика, а затем может применить одну или несколько категорий к любому элементу сообщения или встрече.</span><span class="sxs-lookup"><span data-stu-id="5971f-105">The user defines the master list of color-coded categories for their mailbox, and can then apply one or more of those categories to any message or appointment item.</span></span> <span data-ttu-id="5971f-106">Каждая [Категория](/javascript/api/outlook/office.categorydetails) в главном списке представлена именем и [цветом](/javascript/api/outlook/office.mailboxenums.categorycolor) , указанными пользователем.</span><span class="sxs-lookup"><span data-stu-id="5971f-106">Each [category](/javascript/api/outlook/office.categorydetails) in the master list is represented by the name and [color](/javascript/api/outlook/office.mailboxenums.categorycolor) that the user specifies.</span></span> <span data-ttu-id="5971f-107">С помощью API JavaScript для Office можно управлять главным списком категорий в почтовом ящике и категориями, применяемыми к элементу.</span><span class="sxs-lookup"><span data-stu-id="5971f-107">You can use the Office JavaScript API to manage the categories master list on the mailbox and the categories applied to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="5971f-108">Поддержка этой функции появилась в наборе требований 1,8.</span><span class="sxs-lookup"><span data-stu-id="5971f-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="5971f-109">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="5971f-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="manage-categories-in-the-master-list"></a><span data-ttu-id="5971f-110">Управление категориями в главном списке</span><span class="sxs-lookup"><span data-stu-id="5971f-110">Manage categories in the master list</span></span>

<span data-ttu-id="5971f-111">Только категории в главном списке в вашем почтовом ящике доступны для применения к сообщению или встрече.</span><span class="sxs-lookup"><span data-stu-id="5971f-111">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="5971f-112">С помощью API можно добавлять, получать и удалять главные категории.</span><span class="sxs-lookup"><span data-stu-id="5971f-112">You can use the API to add, get, and remove master categories.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5971f-113">Чтобы надстройка управляла главным списком категорий, необходимо задать `Permissions` для узла манифеста значение `ReadWriteMailbox` .</span><span class="sxs-lookup"><span data-stu-id="5971f-113">For the add-in to manage the categories master list, you must set the `Permissions` node in the manifest to `ReadWriteMailbox`.</span></span>

### <a name="add-master-categories"></a><span data-ttu-id="5971f-114">Добавление основных категорий</span><span class="sxs-lookup"><span data-stu-id="5971f-114">Add master categories</span></span>

<span data-ttu-id="5971f-115">В приведенном ниже примере показано, как добавить категорию с именем "срочно!".</span><span class="sxs-lookup"><span data-stu-id="5971f-115">The following example shows how to add a category named "Urgent!"</span></span> <span data-ttu-id="5971f-116">в главный список, вызывая [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) для [Mailbox. мастеркатегориес](/javascript/api/outlook/office.mailbox#mastercategories).</span><span class="sxs-lookup"><span data-stu-id="5971f-116">to the master list by calling [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
var masterCategoriesToAdd = [
    {
        "displayName": "Urgent!",
        "color": Office.MailboxEnums.CategoryColor.Preset0
    }
];

Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-master-categories"></a><span data-ttu-id="5971f-117">Получение основных категорий</span><span class="sxs-lookup"><span data-stu-id="5971f-117">Get master categories</span></span>

<span data-ttu-id="5971f-118">В приведенном ниже примере показано, как получить список категорий, вызвав метод [Async](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) для [Mailbox. мастеркатегориес](/javascript/api/outlook/office.mailbox#mastercategories).</span><span class="sxs-lookup"><span data-stu-id="5971f-118">The following example shows how to get the list of categories by calling [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a><span data-ttu-id="5971f-119">Удаление основных категорий</span><span class="sxs-lookup"><span data-stu-id="5971f-119">Remove master categories</span></span>

<span data-ttu-id="5971f-120">В приведенном ниже примере показано, как удалить категорию с именем "срочно!".</span><span class="sxs-lookup"><span data-stu-id="5971f-120">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="5971f-121">из основного списка, вызывая [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) для [Mailbox. мастеркатегориес](/javascript/api/outlook/office.mailbox#mastercategories).</span><span class="sxs-lookup"><span data-stu-id="5971f-121">from the master list by calling [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
var masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a><span data-ttu-id="5971f-122">Управление категориями в сообщении или встрече</span><span class="sxs-lookup"><span data-stu-id="5971f-122">Manage categories on a message or appointment</span></span>

<span data-ttu-id="5971f-123">С помощью API можно добавлять, получать и удалять категории для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="5971f-123">You can use the API to add, get, and remove categories for a message or appointment item.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5971f-124">Только категории в главном списке в вашем почтовом ящике доступны для применения к сообщению или встрече.</span><span class="sxs-lookup"><span data-stu-id="5971f-124">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="5971f-125">В этом разделе описано, как [управлять категориями в главном списке](#manage-categories-in-the-master-list) для получения дополнительных сведений.</span><span class="sxs-lookup"><span data-stu-id="5971f-125">See the earlier section [Manage categories in the master list](#manage-categories-in-the-master-list) for more information.</span></span>
>
> <span data-ttu-id="5971f-126">В Outlook в Интернете невозможно использовать API для управления категориями в сообщениях в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="5971f-126">In Outlook on the web, you can't use the API to manage categories on a message in Read mode.</span></span>

### <a name="add-categories-to-an-item"></a><span data-ttu-id="5971f-127">Добавление категорий в элемент</span><span class="sxs-lookup"><span data-stu-id="5971f-127">Add categories to an item</span></span>

<span data-ttu-id="5971f-128">В приведенном ниже примере показано, как применить категорию с именем "срочно!".</span><span class="sxs-lookup"><span data-stu-id="5971f-128">The following example shows how to apply the category named "Urgent!"</span></span> <span data-ttu-id="5971f-129">к текущему элементу, вызывая [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) для `item.categories` .</span><span class="sxs-lookup"><span data-stu-id="5971f-129">to the current item by calling [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) on `item.categories`.</span></span>

```js
var categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-an-items-categories"></a><span data-ttu-id="5971f-130">Получение категорий элемента</span><span class="sxs-lookup"><span data-stu-id="5971f-130">Get an item's categories</span></span>

<span data-ttu-id="5971f-131">В приведенном ниже примере показано, как получить категории, примененные к текущему элементу, вызвав метод [Async](/javascript/api/outlook/office.categories#getasync-options--callback-) `item.categories` .</span><span class="sxs-lookup"><span data-stu-id="5971f-131">The following example shows how to get the categories applied to the current item by calling [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) on `item.categories`.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a><span data-ttu-id="5971f-132">Удаление категорий из элемента</span><span class="sxs-lookup"><span data-stu-id="5971f-132">Remove categories from an item</span></span>

<span data-ttu-id="5971f-133">В приведенном ниже примере показано, как удалить категорию с именем "срочно!".</span><span class="sxs-lookup"><span data-stu-id="5971f-133">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="5971f-134">из текущего элемента, вызывая [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) для `item.categories` .</span><span class="sxs-lookup"><span data-stu-id="5971f-134">from the current item by calling [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) on `item.categories`.</span></span>

```js
var categoriesToRemove = ["Urgent!"];

Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="see-also"></a><span data-ttu-id="5971f-135">См. также</span><span class="sxs-lookup"><span data-stu-id="5971f-135">See also</span></span>

- [<span data-ttu-id="5971f-136">Разрешения Outlook</span><span class="sxs-lookup"><span data-stu-id="5971f-136">Outlook permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="5971f-137">Элемент Permissions в манифесте</span><span class="sxs-lookup"><span data-stu-id="5971f-137">Permissions element in the manifest</span></span>](../reference/manifest/permissions.md)
