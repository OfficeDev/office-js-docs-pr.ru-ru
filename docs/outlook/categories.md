---
title: Получение и Настройка категорий
description: Как управлять категориями для почтового ящика и элемента
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d0bb2e9f51675c263d0a3a130c64e02e7d55b764
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42721025"
---
# <a name="get-and-set-categories"></a>Получение и Настройка категорий

В Outlook пользователь может применять категории к сообщениям и встречам в виде средств Организации их данных почтового ящика. Пользователь определяет главный список категорий для своего почтового ящика, а затем может применить одну или несколько категорий к любому элементу сообщения или встрече. Каждая [Категория](/javascript/api/outlook/office.categorydetails) в главном списке представлена именем и [цветом](/javascript/api/outlook/office.mailboxenums.categorycolor) , указанными пользователем. С помощью API JavaScript для Office можно управлять главным списком категорий в почтовом ящике и категориями, применяемыми к элементу.

> [!NOTE]
> Поддержка этой функции появилась в наборе требований 1,8. См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="manage-categories-in-the-master-list"></a>Управление категориями в главном списке

Только категории в главном списке в вашем почтовом ящике доступны для применения к сообщению или встрече. С помощью API можно добавлять, получать и удалять главные категории.

> [!IMPORTANT]
> Чтобы надстройка управляла главным списком категорий, необходимо задать для `Permissions` `ReadWriteMailbox`узла манифеста значение.

### <a name="add-master-categories"></a>Добавление основных категорий

В приведенном ниже примере показано, как добавить категорию с именем "срочно!". в главный список, вызывая [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) для [Mailbox. мастеркатегориес](/javascript/api/outlook/office.mailbox#mastercategories).

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

### <a name="get-master-categories"></a>Получение основных категорий

В приведенном ниже примере показано, как получить список категорий, вызвав метод [Async](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) для [Mailbox. мастеркатегориес](/javascript/api/outlook/office.mailbox#mastercategories).

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

### <a name="remove-master-categories"></a>Удаление основных категорий

В приведенном ниже примере показано, как удалить категорию с именем "срочно!". из основного списка, вызывая [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) для [Mailbox. мастеркатегориес](/javascript/api/outlook/office.mailbox#mastercategories).

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

## <a name="manage-categories-on-a-message-or-appointment"></a>Управление категориями в сообщении или встрече

С помощью API можно добавлять, получать и удалять категории для элемента сообщения или встречи.

> [!IMPORTANT]
> Только категории в главном списке в вашем почтовом ящике доступны для применения к сообщению или встрече. В этом разделе описано, как [управлять категориями в главном списке](#manage-categories-in-the-master-list) для получения дополнительных сведений.
>
> В Outlook в Интернете невозможно использовать API для управления категориями в сообщениях в режиме чтения.

### <a name="add-categories-to-an-item"></a>Добавление категорий в элемент

В приведенном ниже примере показано, как применить категорию с именем "срочно!". к текущему элементу, [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) вызывая `item.categories`addAsync для.

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

### <a name="get-an-items-categories"></a>Получение категорий элемента

В приведенном ниже примере показано, как получить категории, примененные к текущему [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) элементу, `item.categories`вызвав метод async.

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

### <a name="remove-categories-from-an-item"></a>Удаление категорий из элемента

В приведенном ниже примере показано, как удалить категорию с именем "срочно!". из текущего элемента, вызывая [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) для `item.categories`.

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

## <a name="see-also"></a>См. также

- [Разрешения Outlook](understanding-outlook-add-in-permissions.md)
- [Элемент Permissions в манифесте](../reference/manifest/permissions.md)
